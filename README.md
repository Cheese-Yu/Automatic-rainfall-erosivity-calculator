# 降雨侵蚀力自动计算器
![](https://github.com/Cheese-Yu/Automatic-rainfall-erosivity-calculator/raw/master/r1.jpg)
### 简介
毕业设计做了个简易的降雨侵蚀力计算器，实现I30（30分钟雨强的）的自动筛选，从而计算出降雨侵蚀力。
### 相关公式
时间间隔：`= 本行时间 - 前一行时间` 如果为负数则设置成"#VALUE!"<br/>
水位：`= MID(短信,7,4)`<br/>
雨量 mm：`= 本行水位 - 前一行水位`<br/>
雨强 cm/h：`= 雨量 / 时间间隔 * 6`<br/>
单位雨强动能  J/m^2：`= 210.3 + 89 * LOG(雨强)`<br/>
时段雨强动能 J/m^2：`= 雨量 / 10 * 单位雨强`<br/>
E总（这一时段的总降雨动能）J/m^2：`= 这一时段的所有时段雨强之和 - 这一时段第一个时段雨强`<br/>
I30（最大30分钟雨强）cm/h：`= 这一时段中最大的30分钟雨强`<br/>
R（降雨侵蚀力）：`= E总 * I30 / 100`
### 使用方法
1. 整理好数据格式
2. 打开软件，点击选择文件，选择要计算的数据文件
3. 点击开始计算，等待计算（1500行数据计算时间为1分钟左右，请耐心等待…）
4. 软件界面显示"计算完成"时，点击生成文件即可保存计算结果
### 文件格式
1. 第一行为列名，必须有且不能为空
2. 时间的格式必须为"小时:分钟"，请不要有日期和汉字
3. 不能有空行，不能有空行，不能有空行<br/>
### 判断依据及注意事项
* 两场雨的判别依据<br/>
1. 00:00作为分隔点：即前一天23:59下的雨和后一天00:00的雨为两场雨<br/>
2. 两个时间点的时间间隔超过6小时判为新一场降雨<br/>
* 数据的填写位置<br/>
E总、I30、R数据所在的行,为该时段降雨数据的最后一行<br/>
* 其他<br/>
文件格式请看[exampleFile.xlsx](https://github.com/Cheese-Yu/Automatic-rainfall-erosivity-calculator/blob/master/exampleFile.xlsx)<br/>
计算结果请看[resultFile.xlsx](https://github.com/Cheese-Yu/Automatic-rainfall-erosivity-calculator/blob/master/resultFile.xlsx)<br/>
运行需要安装.NET Framework 4.5
### Contact
Email：[yhyclelo@sina.com](mailto:yhyclelo@sina.com)</br>
