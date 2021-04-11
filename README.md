# 一、简介

本项目主要使用python脚本来实现对Excel文本格式的处理和校对工作，实现了连接Excel，处理Excel中我们需求的逻辑问题，对错误部分进行高亮显示，并做了批注，并且每次运行前都会进行更新，有效地提高了文本处理的效率，最后封装成了可执行的.exe文件，用户可以直接调用而不需要下载任何依赖。

# 二、文本校对逻辑

Excel表格第一行分别是：A1：T4 ID；B1：ID；C1：HL Requirements Template；D1：Modified Comment；E1：Modified HLR；F1：Object Type；G1：Core_SoF_Coverage；H1：Assumption；I1：Derived；J1：Rationale；K1：Additional  Info。后面的每一行都相当于是一个实际用例，具体见Excel表格。

其中，我们需要用的列为E，F，I，J，L。以行来遍历，首先判断第F列的值，该列的值只能是requirement或者comment，①如果是requirements，那么首先计算E列中shall的个数，如果shall的个数不是1或者E列为空，那么填充该单元格颜色并进行批注；然后判断I列的值是否是yes或no，若I列的值是yes，则J列的值不能为空，若I列值为no，则J列的值必须为空。②如果F列的值是comment，首先计算E列中shall的个数，E列不能有shall，若E列的shall个数不为0或者E列为空则填充该单元格颜色并进行批注；然后判断I列的值和J列的值是否为空。③若F列的值既不是requirement也不是comment那么就整行填充颜色并加批注。其中L列主要用来存放E列中shall的个数。

本文代码提供了大量的注释，通俗易懂，见excel.py

# 三、本文用到的库

1：openpyxl 是一个用于处理xlsx格式Excel表格文件的第三方库

2：from openpyxl.styles import PatternFill 主要用于导入单元格的填充模块

3：from openpyxl.comments import Comment 主要用于导入批注模块

4：PyInstaller  将Python程序打包成exe可执行文件，该打包方式会增加两个目录build和dist，dist下面的文件就是可以发布的可执行文件，对于上面的命令你会发现dist目录下面有一堆文件，各种都动态库文件和可执行文件。

# 四、需要注意的点

1：python连接Excel，注意Excel中的单元格是从1行1列开始的，而不是0

2：在每次执行前需要关闭我们的Excel表格

