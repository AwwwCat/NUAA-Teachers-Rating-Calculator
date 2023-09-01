# Teachers-Rating-Calculator
南京航空航天大学航天学院教学科研岗(业务型)2020-2022聘期任务基准要求分数计算程序。
# 使用方法
需要 **Visual Studio** 打开该项目。若没有请前往[官网](https://visualstudio.microsoft.com/ "visual studio官网")下载。  
该项目为 C#控制台窗口 项目，请确保vs里下载了 `C# 控制台应用` 模板。  
使用vs打开`Teachers Scores.sln`文件即可。  
需要调整的代码请**参考代码注释**。
# 考核方式
* 一般教师  
![](https://github.com/AwwwCat/Teachers-Rating-Calculator/blob/main/Readme/1.jpg)
![](https://github.com/AwwwCat/Teachers-Rating-Calculator/blob/main/Readme/2.jpg)  
* 实验岗教师
![](https://github.com/AwwwCat/Teachers-Rating-Calculator/blob/main/Readme/3.jpg)  
# 文件说明
所有文件均应为两列，第一列为姓名，第二列为数据，不得包含其他数据或标注（例如：表头、总计等）  
（教师姓名数据若导入成功有其他标注也无所谓==）。  
## CSV文件夹
该文件夹包含了输出的CSV文件  
文件(.csv) | 说明
:---: | :---:
page1 | 程序输出文件，包含最终的考核结果
## Data文件夹
储存输入数据的文件夹，以下excel文件第一列均为教师名称，第二列为数据
文件(.xlsx) | 说明 | 第二列数据录入类型
:---: | :---: | :---:
S0 | 教师岗位等级 | int
S0-1 | 教师科研岗位 | string
S1-1-1 | 本科生课时（不含毕业设计） | double
S1-1-2 | 本科生毕业设计课时 | double
S1-1-3 | 班主任课时 | double
S1-2 | 研究生课时 | double
S1-3 | 本科生导师课时 | double
S1-4 | 指导科创课时 | double
S2-1 | 国家级新增立项项目 | double
S2-2 | 省部级新增立项项目 | double
S2-3 | 其他新增立项项目 | double
S3 | 三年科研到款 | double
aS4-1至aS4-11 | 分别对应二三四岗教师S4的11项 | double
bS4-1至bS4-12 | 分别对应五六七岗教师S4的12项 | double
cS4-4至cS4-10 | 分别对应八九十岗教师S4的第4至10项 | double
S5-1-1 | SCI期刊论文 | double
S5-1-2 | EI期刊论文 | double
S5-1-3 | 重要核心期刊论文 | double
S5-2 | 专利成果转化 | double
S5-3 | 国防科技成果鉴定的成果 | double
S5-4 | 科研成果奖 | double
aS5-5 | 二三四岗教师新增主持国家重大项目 | double
bS5-5 | 五六七岗教师新增主持国家重大项目 | double
cS5-5 | 八九十岗教师新增主持国家重大项目 | double
S5-6 | 省部级以上科研平台 | double
dS2-1至dS2-12 | 分别对应实验岗五六七岗教师S2的前12项 | double
eS2-1至eS2-7 | 分别对应实验岗八九十岗教师S2的前7项 | double
dS3 | 实验岗教师实验室工作完成 | double
S6 | 集体工作达到要求 | double
## TXT文件夹
该文件夹包含了输出的TXT文件
## Teachers Scores文件夹
该文件夹包含了项目文件
