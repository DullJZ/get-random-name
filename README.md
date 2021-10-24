# 点名程序使用文档

https://dulljzblog.jz-home.top/2021/10/24/random-name-use/


## 特点

1. 支持读取学生名单
2. 支持加权点名
3. 调用一言API

## 软件截图

<img src="https://cdn.jsdelivr.net/gh/DullJZ/MyPicture/点名程序截图.png" style="zoom:50%;" />



## 功能

### 读取学生名单

软件运行时会自动对同目录下的 `students.txt` 进行读取并显示在列表框中。

<img src="https://cdn.jsdelivr.net/gh/DullJZ/MyPicture/random-name-load-students.png" style="zoom:50%;" />

要导入学生信息，只需修改 `students.txt` 即可。

**每行一人，按学号编排，不可空行**

**以GB2312方式编码，否则出现乱码问题！**

<img src="https://cdn.jsdelivr.net/gh/DullJZ/MyPicture/random-name-encoding.png" style="zoom:50%;" />



### 加权点名

加权算法：以10000为基数，乘以权重，得到该同学的姓名重复次数(n)；其余(10000-n)个为随机生成的学生姓名；最后在10000个姓名中随机抽取

其中：学号为导入的学生顺序号

### 一言API

目前使用的一言API有：

https://v1.hitokoto.cn/?encode=text

https://api.btstu.cn/yan/api.php

https://api.muxiaoguo.cn/api/dujitang

本地导入一言库正在开发中！

**一言内容与作者观点无关！**