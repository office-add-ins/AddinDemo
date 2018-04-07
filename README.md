
## 环境
- [x]  VS2017 15.5.1
- [x]  Office 2016 professional plus (16.0.9029.2253)


## Add-in问题一览

-  1、企业版pro plus，无法升级，导致无法加载Add-in Commands。（**非常紧急**）截图如下

-  2、New Word JavaScript APIs (1.4)，添加了CreateDocument方法，但是不能对create后的Document做任何操作。（**非常紧急**）

-  3、office.js暂不支持插入OLE对象，如Excel等（紧急）

-  4、无法定位shape，及嵌套在shape的table。

-  5、add-in commands，一个group为何只能放2列？

-  6、 command文字，为何4个中文字，中间会换行？，如（我爱你们，目前显示在2行，希望显示在一行）

-  7、目前Office UI库，有以下2个cdn，并且两者使用的样式，方法 有很大区别。
      只有jquery能在vs里调试，react无法在vs里调试，

-  8、 无法检测到Word或Excel等关闭，打开等事件。
      
## 参考图
**问题1**
<img src="https://raw.githubusercontent.com/office-add-ins/AddinDemo/master/version1.png"  alt="图片描述文字"/>
<img src="https://raw.githubusercontent.com/office-add-ins/AddinDemo/master/version2.png"  alt="图片描述文字"/>


**问题5、6**
<img src="https://raw.githubusercontent.com/office-add-ins/AddinDemo/master/menu.png"  alt="图片描述文字"/>
