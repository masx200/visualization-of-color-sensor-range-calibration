"""这段代码的功能是将给定的颜色数据绘制在3D色彩空间中，以直观地展示不同颜色的色彩空间比值范围。它首先从excel_data列表中获取颜色的数据，并将其存储在colors_data字典中，然后使用matplotlib库中的3D绘图功能创建一个3D坐标系，根据colors_data中的色彩空间比值范围为每个颜色创建一个多边形区域，并用相应的颜色填充。最后，它显示了包含所有颜色多边形区域的3D图形。
    """
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
import math

excel_data = [   # 数据列表，每个元素是一个字典
 
{
"积木照片": '=DISPIMG("ID_1725801A2D224C488B9ACAAC13DF5581",1)',
"颜色名": "黑",
"颜色代号": "D",
"颜色色值": "#000000",
"R/G下限": "1",
"R/G上限": "1",
"R/B下限": "1",
"R/B上限": "1.4",
"G/B下限": "1",
"G/B上限": "1.4",
"R/G理论值": "1",
"R/B理论值": "1",
"G/B理论值": "1",
},
{
"积木照片": '=DISPIMG("ID_28B20C772D3F4DE080F8A36E3465D9E8",1)',
"颜色名": "白",
"颜色代号": "W",
"颜色色值": "#FFFFFF",
"R/G下限": "1.071428571",
"R/G上限": "1.128205128",
"R/B下限": "1.452830189",
"R/B上限": "1.666666667",
"G/B下限": "1.339622642",
"G/B上限": "1.555555556",
"R/G理论值": "1",
"R/B理论值": "1",
"G/B理论值": "1",
},
{
"积木照片": '=DISPIMG("ID_F65C608A77C849F99BF5918CE3CD2D42",1)',
"颜色名": "红",
"颜色代号": "R",
"颜色色值": "#FF0000",
"R/G下限": "2.375",
"R/G上限": "9",
"R/B下限": "3.166666667",
"R/B上限": "9",
"G/B下限": "1",
"G/B上限": "2",
"R/G理论值": "∞",
"R/B理论值": "∞",
"G/B理论值": "1",
},
{
"积木照片": '=DISPIMG("ID_C65E8798A6FF4BD7910797A657BB9820",1)',
"颜色名": "黄",
"颜色代号": "Y",
"颜色色值": "#FFFF00",
"R/G下限": "1.083333333",
"R/G上限": "1.888888889",
"R/B下限": "2.565217391",
"R/B上限": "5.666666667",
"G/B下限": "2.142857143",
"G/B上限": "3.333333333",
"R/G理论值": "1",
"R/B理论值": "∞",
"G/B理论值": "∞",
},
{
"积木照片": '=DISPIMG("ID_71B15858F5BD408B964BEC7EC4BDC05B",1)',
"颜色名": "绿",
"颜色代号": "G",
"颜色色值": "#00FF00",
"R/G下限": "0.1",
"R/G上限": "0.6",
"R/B下限": "0.666666667",
"R/B上限": "1.5",
"G/B下限": "1.5",
"G/B上限": "10",
"R/G理论值": "0",
"R/B理论值": "1",
"G/B理论值": "∞",
},
{
"积木照片": '=DISPIMG("ID_AE86ED07E4194926A955AFBFF9A7EEF1",1)',
"颜色名": "青",
"颜色代号": "C",
"颜色色值": "#00FFFF",
"R/G下限": "0.333333333",
"R/G上限": "0.936170213",
"R/B下限": "0.333333333",
"R/B上限": "1.111111111",
"G/B下限": "0.833333333",
"G/B上限": "1.230769231",
"R/G理论值": "0",
"R/B理论值": "0",
"G/B理论值": "1",
},
{
"积木照片": '=DISPIMG("ID_D7EB52323C2F491ABBF4765B59104844",1)',
"颜色名": "蓝",
"颜色代号": "B",
"颜色色值": "#0000FF",
"R/G下限": "0.5",
"R/G上限": "1",
"R/B下限": "0",
"R/B上限": "0.571428571",
"G/B下限": "0.333333333",
"G/B上限": "0.6",
"R/G理论值": "1",
"R/B理论值": "0",
"G/B理论值": "0",
},
{
"积木照片": '=DISPIMG("ID_E171769754DE43E2ABF2EDE78E1F0BB2",1)',
"颜色名": "紫",
"颜色代号": "P",
"颜色色值": "#800080",
"R/G下限": "1",
"R/G上限": "1E+308",
"R/B下限": "0.5",
"R/B上限": "2",
"G/B下限": "0",
"G/B上限": "1",
"R/G理论值": "∞",
"R/B理论值": "1",
"G/B理论值": "0",
},
]

colors_data = {  # 颜色数据字典，键是颜色代号，值是包含色彩空间比值范围和颜色的字典
    
    "D": {"limits": [(1, 1.34), (1, 1.5), (1, 1.5)], "color": "#000000"},
"W": {"limits": [(1.07, 1.13), (1.32, 1.5), (1.14, 1.29)], "color": "#FFFFFF"},
"R": {"limits": [(3.5, 7.2), (4.33, 9), (1, 2)], "color": "#FF0000"},
"Y": {"limits": [(1.47, 1.6), (3.35, 5.25), (2.16, 2.75)], "color": "#FFFF00"},
"G": {"limits": [(0.25, 0.58), (0.5, 1), (1.5, 2)], "color": "#00FF00"},
"C": {"limits": [(0.28, 0.6), (0.25, 0.6), (0.75, 1)], "color": "#00FFFF"},
"B": {"limits": [(0.33, 0.67), (0.16, 0.4), (0.5, 0.66)], "color": "#0000FF"},
"P": {"limits": [(1.3, 2), (0.91, 1.08), (0.5, 0.77)], "color": "#800080"},

}

for data in excel_data:  # 遍历excel_data中的每一个字典
    colors_data[data["颜色代号"]] = {
        "limits": [
            (
                math.log10(max(0.1, min(10, float(data["R/G下限"])))),
                math.log10(max(0.1, min(10, float(data["R/G上限"])))),
            ),
            (
                math.log10(max(0.1, min(10, float(data["R/B下限"])))),
                math.log10(max(0.1, min(10, float(data["R/B上限"])))),
            ),
            (
                math.log10(max(0.1, min(10, float(data["G/B下限"])))),
                math.log10(max(0.1, min(10, float(data["G/B上限"])))),
            ),
        ],
        "color": data["颜色色值"],
    }
# 创建3D坐标的画布
fig = plt.figure()
ax = fig.add_subplot(111, projection="3d")

# 根据色彩空间比值范围绘制每一个颜色矩形区域
for color_name, color_info in colors_data.items():
    # 获取色彩空间比值范围和颜色
    limits = color_info["limits"]
    color = color_info["color"]
    rg_limits, rb_limits, gb_limits = limits
    vertices = [   # 顶点列表，用于创建多边形
        [rg_limits[0], rb_limits[0], gb_limits[0]],
        [rg_limits[1], rb_limits[0], gb_limits[0]],
        [rg_limits[1], rb_limits[1], gb_limits[0]],
        [rg_limits[0], rb_limits[1], gb_limits[0]],
        [rg_limits[0], rb_limits[0], gb_limits[1]],
        [rg_limits[1], rb_limits[0], gb_limits[1]],
        [rg_limits[1], rb_limits[1], gb_limits[1]],
        [rg_limits[0], rb_limits[1], gb_limits[1]],
    ]
    faces = [   # 多边形面列表，用于创建多边形集合
        [vertices[j] for j in [0, 1, 2, 3]],
        [vertices[j] for j in [4, 5, 6, 7]],
        [vertices[j] for j in [0, 1, 5, 4]],
        [vertices[j] for j in [2, 3, 7, 6]],
        [vertices[j] for j in [1, 2, 6, 5]],
        [vertices[j] for j in [0, 3, 7, 4]],
    ]

    # 创建并添加3D多边形集合
    ax.add_collection3d(
        Poly3DCollection(
            faces, facecolors=color, linewidths=1, edgecolors="k", alpha=0.5
        )
    )

    # 添加每个颜色的名字标签
    mid_x = (rg_limits[0] + rg_limits[1]) / 2
    mid_y = (rb_limits[0] + rb_limits[1]) / 2
    mid_z = (gb_limits[0] + gb_limits[1]) / 2
    ax.text(mid_x, mid_y, mid_z, color_name, color="black")

# 标记坐标轴
ax.set_xlabel("lg(R/G Axis)")
ax.set_ylabel("lg(R/B Axis)")
ax.set_zlabel("lg(G/B Axis)")

# 设置图形的显示范围以便查看所有颜色空间矩形
ax.set_xlim(-1.0, 1.0)
ax.set_ylim(-1.0, 1.0)
ax.set_zlim(-1.0, 1.0)

# 显示3D plot
plt.show()
