## 基于Python自动点名脚本
### 一.安装相关库
安装pptx库

``` pip install python-pptx ```
### 二.使用方法
将需要点名的照片以``` 姓名+jpg ```的格式命名放在主文件目录的``` picture ```文件夹下

最后保存即可

### 三.注意事项
在使用程序后生成的pptx需要手动更改一些内容

##### 1.将自动换页选项打开，设置换页时间为0.1s，重复，并应用于全部幻灯片
##### 2.您可以对幻灯片样式进行个性化设置
##### 3.如果您想对图片大小及文本框字体及相关设置，请对源码进行更改

以下是对一些自定义操作的：

（1）图片位置设置

```
#图片位置设置
leftimg = util.Cm(10.61)
topimg = util.Cm(1.39)
widthimg = util.Cm(11)
heightimg = util.Cm(11)
```

其中四项代表了图片在幻灯片的相对位置

（2）文本框设置

```
#文本框位置设置
lefttext = util.Cm(11.30)
toptext = util.Cm(12.94)
widthtext = util.Cm(9.39)
heighttext = util.Cm(4)
```
对应上信息进行更改即可