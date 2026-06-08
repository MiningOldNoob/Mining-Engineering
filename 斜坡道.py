# -*- coding: utf-8 -*-
# 指定 Python 文件编码为 UTF-8，防止中文注释在部分环境下乱码。

"""
按图示逻辑生成三维折返斜坡道中心线：
跑道形折返单元 + 层间平面错移 + 高程递增

依赖：
pip install pywin32
"""
# 多行字符串说明本脚本的用途、几何逻辑和运行依赖。

import math
# 导入 math 数学库，用于三角函数、弧度转换、圆周率 pi、向上取整等计算。

import pythoncom
# 导入 pythoncom，用于构造 AutoCAD COM 接口所需的数据类型。

import win32com.client
# 导入 win32com.client，用于通过 Python 调用 AutoCAD 软件接口。

from win32com.client import VARIANT

# 导入 VARIANT 类型，用于把 Python 数组转换为 AutoCAD 可识别的数组格式。


# =========================================================
# 1. 参数区
# =========================================================
# 以下为建模参数，可以根据实际斜坡道设计要求修改。


# 工程起点坐标
X0 = 0.0
# 斜坡道中心线起点的工程 X 坐标。

Y0 = 0.0
# 斜坡道中心线起点的工程 Y 坐标。

Z0 = 0.0
# 斜坡道中心线起点的工程 Z 坐标，即起点高程。


# 图示参数
L = 32.25
# 跑道形折返单元中两侧弯道切点之间的直线段长度。

R = 12
# 折返弯道中心线半径”。


# 坡度参数
i_straight = 0.12
# 直线段坡度，按小数表示；0.12 表示 12%。

i_curve = 0.03
# 弯道段坡度，按小数表示；0.03 表示 3%。


# 循环层数
n_loop = 6
# 需要生成的折返循环数量；数值越大，生成的斜坡道层数越多。


# 弯道离散边长
arc_step = 0.5
# 用多段线近似圆弧时，每个弧线小段的目标长度；越小越圆滑，但节点越多。


# 整体平面错移方向，单位：度
# 以 CAD X 正方向为 0°，逆时针为正
advance_azimuth_deg = 90
# 每完成一个折返单元后，整体向该方位角方向发生平面错移。


# 每完成一个完整折返单元后的平面错移距离
# 该值控制多层之间在平面上的错开程度
advance_distance = 1.0
# 每一循环相对于上一循环的平面偏移距离；该值越大，多层之间错开越明显。


# 整体旋转角
# 若希望跑道形单元自身再旋转，可修改该参数
base_rotation_deg = 0.0
# 控制每一个跑道形折返单元自身的整体旋转角度。


# 是否标注节点
ADD_MARKERS = True


# True 表示在 CAD 中添加关键节点文字和点标记；False 表示只画中心线。


# =========================================================
# 2. 工具函数
# =========================================================
# 以下函数用于坐标旋转、坐标格式转换和 AutoCAD 数据封装。


def rotate_xy(x, y, angle_deg):
    # 定义二维坐标旋转函数，用于将局部平面坐标旋转到指定角度。

    """二维旋转"""
    # 函数说明：输入 x、y 和旋转角度，输出旋转后的坐标。

    a = math.radians(angle_deg)
    # 将角度制转换为弧度制，因为 Python 三角函数使用弧度。

    xr = x * math.cos(a) - y * math.sin(a)
    # 二维旋转公式，计算旋转后的 X 坐标。

    yr = x * math.sin(a) + y * math.cos(a)
    # 二维旋转公式，计算旋转后的 Y 坐标。

    return xr, yr
    # 返回旋转后的二维坐标。


def to_variant_point(x, y, z):
    # 定义将单个三维点转换为 AutoCAD COM 接口可识别格式的函数。

    return VARIANT(
        # 返回 VARIANT 类型对象，这是 AutoCAD COM 接口要求的数据格式。

        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        # 指定数据类型为双精度浮点数组。

        [float(x), float(y), float(z)]
        # 将点坐标转换为浮点型列表，格式为 [X, Y, Z]。
    )


def make_variant_coords(points):
    # 定义将多个三维点转换为 AutoCAD 3DPolyline 坐标数组的函数。

    coords = []
    # 创建空列表，用于存储一维坐标序列。

    for x, y, z in points:
        # 遍历所有三维点。

        coords.extend([float(x), float(y), float(z)])
        # 将每个点的 X、Y、Z 依次展开，形成 [x1,y1,z1,x2,y2,z2,...] 格式。

    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, coords)
    # 将一维坐标数组转换为 AutoCAD 可读取的 VARIANT 双精度数组。


# =========================================================
# 3. 生成图示逻辑下的中心线节点
# =========================================================
# 以下函数是几何计算核心，用于生成三维斜坡道中心线节点。


def generate_racetrack_ramp_points(
        # 定义生成跑道形折返斜坡道中心线节点的函数。

        X0, Y0, Z0,
        # 输入工程起点坐标。

        L, R,
        # 输入直线段长度 L 和弯道半径 R。

        i_straight, i_curve,
        # 输入直线段坡度和弯道段坡度。

        n_loop,
        # 输入循环层数。

        arc_step,
        # 输入弯道离散步长。

        advance_azimuth_deg,
        # 输入整体平面错移方向。

        advance_distance,
        # 输入每循环平面错移距离。

        base_rotation_deg
        # 输入跑道形单元自身旋转角。
):
    """
    生成图示逻辑的三维多段线节点。

    一个完整循环：
    1）下侧直线：左 → 右
    2）右侧 180°弯道：下 → 上
    3）上侧直线：右 → 左
    4）左侧 180°弯道：上 → 下

    与普通闭合跑道不同：
    沿整个循环累计过程，中心线逐步叠加一个平面错移量，
    使循环终点不回到原点，而是进入下一循环起点。
    """
    # 函数说明：该函数生成非闭合的三维折返斜坡道中心线。

    # 弯道分段数
    m = max(4, math.ceil(math.pi * R / arc_step))
    # 计算单个 180°弯道的离散段数。
    # math.pi * R 是半圆弧长。
    # arc_step 是目标离散边长。
    # math.ceil 表示向上取整。
    # max(4, ...) 保证半圆至少分成 4 段，避免过于粗糙。

    # 一个循环的水平展开长度
    loop_len = 2.0 * L + 2.0 * math.pi * R
    # 一个完整循环包括两段直线和两个 180°弯道。
    # 两段直线长度为 2L。
    # 两个半圆弧合计长度为 2πR。

    # 一个循环的垂高
    h_loop = 2.0 * i_straight * L + 2.0 * i_curve * math.pi * R
    # 计算每完成一个完整循环后的高程增量。
    # 直线段垂高为 i_straight × L，两段直线合计 2*i_straight*L。
    # 弯道段垂高为 i_curve × πR，两个弯道合计 2*i_curve*πR。

    # 层间平面错移向量
    beta = math.radians(advance_azimuth_deg)
    # 将整体错移方向由角度转换为弧度。

    dx_adv = advance_distance * math.cos(beta)
    # 计算每循环平面错移在 X 方向上的分量。

    dy_adv = advance_distance * math.sin(beta)
    # 计算每循环平面错移在 Y 方向上的分量。

    points = []
    # 创建列表，用于存储中心线所有三维节点。

    key_points = []

    # 创建列表，用于存储关键节点及其名称，例如 A0、B0、C0、D0。

    def add_point(k, x_local, y_local, z_add, s_cum, name=None):
        # 定义内部函数，用于添加单个中心线节点。
        # k 表示当前循环编号。
        # x_local、y_local 表示跑道形单元中的局部平面坐标。
        # z_add 表示当前点在本循环内的相对高程增量。
        # s_cum 表示当前点在本循环内的累计水平展开长度。
        # name 表示关键节点名称，可为空。

        """
        k: 当前循环编号
        s_cum: 当前点在本循环内的累计水平展开长度
        """
        # 内部函数说明。

        u = s_cum / loop_len
        # 计算当前点在本循环中的进度比例。
        # u=0 表示循环起点。
        # u=1 表示循环终点。

        # 层间错移：k 个完整循环 + 当前循环内的渐进错移
        dx = (k + u) * dx_adv
        # 计算当前点累计平面错移的 X 分量。
        # k 表示已经完成的循环数。
        # u 表示当前循环内的渐进错移比例。

        dy = (k + u) * dy_adv
        # 计算当前点累计平面错移的 Y 分量。

        # 先对跑道单元自身旋转
        xr, yr = rotate_xy(x_local, y_local, base_rotation_deg)
        # 将局部跑道形单元按照 base_rotation_deg 进行整体旋转。

        X = X0 + xr + dx
        # 计算工程 X 坐标：起点 X0 + 旋转后的局部 X + 平面错移 X。

        Y = Y0 + yr + dy
        # 计算工程 Y 坐标：起点 Y0 + 旋转后的局部 Y + 平面错移 Y。

        Z = Z0 + k * h_loop + z_add
        # 计算工程 Z 坐标：起点高程 + 已完成循环高程 + 当前循环内高程增量。

        p = (X, Y, Z)
        # 将 X、Y、Z 组合为一个三维点。

        points.append(p)
        # 将该点加入中心线节点列表。

        if name:
            # 如果当前点设置了关键节点名称，则执行记录。

            key_points.append((name, p))
            # 将关键节点名称和坐标一起保存，便于后续在 CAD 中标注。

    for k in range(n_loop):
        # 按循环层数依次生成每一个折返单元。
        # k 从 0 到 n_loop-1。

        # ---------------------------
        # A：下侧左切点
        # ---------------------------
        # A 点为当前折返单元的起点。

        s_cum = 0.0
        # A 点位于当前循环的开始位置，累计水平展开长度为 0。

        z_add = 0.0
        # A 点在当前循环内的相对高程增量为 0。

        if k == 0:
            # 只在第一个循环时显式添加 A0。
            # 后续循环的 A 点已经由前一循环的终点生成，避免重复点。

            add_point(k, 0.0, 0.0, z_add, s_cum, f"A{k}")
            # 添加 A0 点，局部坐标为 (0,0)，并标记名称为 A0。

        # ---------------------------
        # B：下侧右切点
        # ---------------------------
        # B 点为下侧直线段终点，也是右侧弯道起点。

        s_cum = L
        # 从 A 到 B 的累计水平展开长度为 L。

        z_add = i_straight * L
        # 下侧直线段高程增量为坡度 × 水平长度。

        add_point(k, L, 0.0, z_add, s_cum, f"B{k}")
        # 添加 B 点，局部坐标为 (L,0)。

        # ---------------------------
        # 右侧 180°弯道
        # 圆心：(L, R)
        # t = 0：下切点
        # t = pi：上切点
        # ---------------------------
        # 该段由 B 点向上转弯至 C 点。

        for j in range(1, m + 1):
            # 按 m 个离散段生成右侧半圆弯道节点。
            # 从 1 开始，是为了避免重复添加 B 点。

            t = j * math.pi / m
            # 计算当前离散点对应的圆弧参数 t。
            # t 从接近 0 逐渐增加到 π。

            x = L + R * math.sin(t)
            # 右侧半圆弯道的局部 X 坐标。
            # 当 t=0 时 x=L；当 t=π/2 时 x=L+R；当 t=π 时 x=L。

            y = R - R * math.cos(t)
            # 右侧半圆弯道的局部 Y 坐标。
            # 当 t=0 时 y=0；当 t=π 时 y=2R。

            s_cum = L + R * t
            # 当前点累计水平展开长度 = 下侧直线长度 + 已走过的弯道弧长。

            z_add = i_straight * L + i_curve * R * t
            # 当前点相对高程 = 下侧直线增高 + 右侧弯道已增高。

            name = f"C{k}" if j == m else None
            # 若当前点为右侧弯道最后一个点，则命名为 Ck；否则不命名。

            add_point(k, x, y, z_add, s_cum, name)
            # 将右侧弯道当前离散点加入中心线节点列表。

        # ---------------------------
        # D：上侧左切点
        # 上侧直线：右 → 左
        # ---------------------------
        # D 点为上侧直线段终点，也是左侧弯道起点。

        for q in range(1, 2):
            # 该循环只执行一次。
            # 这里写成 for 循环不是必要的，可直接写赋值和 add_point。
            # 保留该结构不会影响结果。

            x = 0.0
            # D 点局部 X 坐标为 0，位于左侧。

            y = 2.0 * R
            # D 点局部 Y 坐标为 2R，位于上侧直线左端。

            s_cum = L + math.pi * R + L
            # D 点累计水平展开长度 = 下侧直线 L + 右侧半圆 πR + 上侧直线 L。

            z_add = i_straight * L + i_curve * math.pi * R + i_straight * L
            # D 点相对高程 = 下侧直线增高 + 右侧半圆增高 + 上侧直线增高。

            add_point(k, x, y, z_add, s_cum, f"D{k}")
            # 添加 D 点，并标记名称为 Dk。

        # ---------------------------
        # 左侧 180°弯道
        # 圆心：(0, R)
        # t = 0：上切点
        # t = pi：下切点
        # ---------------------------
        # 该段由 D 点向下转弯，形成下一循环起点 A(k+1)。

        for j in range(1, m + 1):
            # 按 m 个离散段生成左侧半圆弯道节点。
            # 从 1 开始，是为了避免重复添加 D 点。

            t = j * math.pi / m
            # 计算左侧弯道当前离散点的圆弧参数。

            x = -R * math.sin(t)
            # 左侧半圆弯道的局部 X 坐标。
            # 当 t=0 时 x=0；当 t=π/2 时 x=-R；当 t=π 时 x=0。

            y = R + R * math.cos(t)
            # 左侧半圆弯道的局部 Y 坐标。
            # 当 t=0 时 y=2R；当 t=π 时 y=0。

            s_cum = L + math.pi * R + L + R * t
            # 当前点累计水平展开长度 = 下侧直线 + 右侧弯道 + 上侧直线 + 左侧弯道已走弧长。

            z_add = (
                    i_straight * L
                    # 下侧直线段产生的高程增量。

                    + i_curve * math.pi * R
                    # 右侧 180°弯道产生的高程增量。

                    + i_straight * L
                    # 上侧直线段产生的高程增量。

                    + i_curve * R * t
                # 左侧弯道当前已产生的高程增量。
            )

            name = f"A{k + 1}" if j == m else None
            # 若当前点为左侧弯道最后一个点，则命名为下一循环起点 A(k+1)。

            add_point(k, x, y, z_add, s_cum, name)
            # 将左侧弯道当前离散点加入中心线节点列表。

    return points, key_points, h_loop, loop_len
    # 返回四个结果：
    # points：所有中心线三维节点。
    # key_points：关键节点及其名称。
    # h_loop：每循环垂高。
    # loop_len：每循环水平展开长度。


# =========================================================
# 4. 调用 AutoCAD 生成 3DPolyline
# =========================================================
# 以下函数负责连接 AutoCAD，并把节点绘制成三维多段线。


def draw_3d_polyline_in_cad(points, key_points=None):
    # 定义绘图函数。
    # points 为三维中心线节点。
    # key_points 为可选关键节点标注数据。

    acad = win32com.client.Dispatch("AutoCAD.Application")
    # 启动或连接当前 AutoCAD 应用程序。

    acad.Visible = True
    # 设置 AutoCAD 界面为可见状态。

    doc = acad.ActiveDocument
    # 获取当前活动的 CAD 图形文件。

    ms = doc.ModelSpace
    # 获取模型空间，后续图形对象都绘制在模型空间中。

    layer_name = "Ramp_Centerline_3D"
    # 定义中心线所在图层名称。

    try:
        # 尝试获取已有图层。

        layer = doc.Layers.Item(layer_name)
        # 如果图层已存在，则直接获取该图层。

    except Exception:
        # 如果图层不存在，则进入异常处理。

        layer = doc.Layers.Add(layer_name)
        # 新建名为 Ramp_Centerline_3D 的图层。

    doc.ActiveLayer = layer
    # 将当前图层设置为斜坡道中心线图层。

    coords = make_variant_coords(points)
    # 将中心线节点转换为 AutoCAD Add3DPoly 可识别的一维坐标数组。

    pline = ms.Add3DPoly(coords)
    # 在模型空间中创建三维多段线对象。

    pline.Layer = layer_name
    # 将三维多段线放置到指定图层。

    pline.Update()
    # 更新三维多段线对象，使其在 CAD 中显示。

    if key_points and ADD_MARKERS:
        # 如果存在关键节点，且设置了允许标注，则添加节点标注。

        for name, p in key_points:
            # 遍历关键节点列表。

            x, y, z = p
            # 拆分关键节点坐标。

            ms.AddPoint(to_variant_point(x, y, z))
            # 在关键节点位置添加 CAD 点对象。

            ms.AddText(name, to_variant_point(x, y, z + 0.5), 0.8)
            # 在关键节点上方 0.5 m 处添加文字标注。
            # 文字高度为 0.8。

    doc.Regen(1)
    # 重新生成图形显示，相当于 CAD 中的 REGEN。

    return acad, doc, pline
    # 返回 AutoCAD 应用对象、当前文档对象和三维多段线对象。


# =========================================================
# 5. 主程序
# =========================================================
# 以下为脚本直接运行时执行的部分。


if __name__ == "__main__":
    # 判断当前文件是否作为主程序运行。
    # 如果是直接运行该 py 文件，则执行以下代码。
    # 如果是被其他 Python 文件导入，则不执行以下代码。

    points, key_points, h_loop, loop_len = generate_racetrack_ramp_points(
        # 调用节点生成函数，返回中心线点、关键点、每循环垂高和水平展开长度。

        X0, Y0, Z0,
        # 输入工程起点坐标。

        L, R,
        # 输入直线长度和弯道半径。

        i_straight, i_curve,
        # 输入直线段坡度和弯道段坡度。

        n_loop,
        # 输入循环层数。

        arc_step,
        # 输入弯道离散步长。

        advance_azimuth_deg,
        # 输入整体平面错移方向。

        advance_distance,
        # 输入每循环平面错移距离。

        base_rotation_deg
        # 输入跑道形单元自身旋转角。
    )

    draw_3d_polyline_in_cad(points, key_points)
    # 调用 AutoCAD 绘图函数，将生成的点连接为三维多段线，并标注关键节点。

    print("三维斜坡道中心线已生成。")
    # 在 Python 控制台输出运行完成提示。

    print(f"直线长度 L = {L:.3f} m")
    # 输出直线段长度，保留三位小数。

    print(f"弯道半径 R = {R:.3f} m")
    # 输出弯道半径，保留三位小数。

    print(f"循环层数 = {n_loop}")
    # 输出循环层数。

    print(f"每循环水平展开长度 = {loop_len:.3f} m")
    # 输出每个循环的中心线水平展开长度。

    print(f"每循环垂高 = {h_loop:.3f} m")
    # 输出每个循环的高程增量。

    print(f"每循环平面错移距离 = {advance_distance:.3f} m")
    # 输出每完成一个循环后的平面错移距离。

    print(f"平面错移方向 = {advance_azimuth_deg:.3f}°")
    # 输出平面错移方向角。

    print(f"总节点数 = {len(points)}")
    # 输出生成的中心线总节点数量。
    # 最终版！