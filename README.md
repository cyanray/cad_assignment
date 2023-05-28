# CAD Assignment

本项目为《船舶CAD与软件开发》课程大作业的 VBA 程序 —— 根据型值表在 AutoCAD 软件中绘制型线图。

## 项目结构

VBA 的工程文件为二进制文件 dvb，不便于查看，因此本仓库将其中的代码导出为文本文件（.vba 或 .cls）。如果有必要，请按照如下结构重新构建你的 VBA 项目。或者直接使用本仓库中的 Project.dvb。

```
Project
├─ AutoCAD对象
│  └─ ThisDrawing
│
├─ 模块(Module)
│  ├─ Utils
│  └─ DataLoader
│
├─ 类模块(ClassModule)
│  ├─ AcadBlockProxy
│  ├─ CurveSpline
│  ├─ Point3
│  ├─ Point3Collection
│  ├─ ShipOffsets
│  └─ Station
│ 
```

型值表位于 **Assets/ShipOFF.txt** 文件中。

绘图结果为 **Assets/result.dwg** 文件（AutoCAD 2007 格式）。

## 其他说明
本项目代码仅供参考。



