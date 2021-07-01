# shanxinglib
各种低效轮子(Various inefficient wheels)

## 框架版本 
.NET 4.8

注：此Package依赖于Newtonsoft.Json，但是我们这个项目用不到，所以不需要也可以，可以在NuGet管理器中勾选 ‘强制卸载，即使有依赖项’，然后强制卸载依赖包Newtonsoft.Json。  
### 第三方库

## 错误
1.无法获取密钥文件“ShanXingTech.pfx”的 MD5 校验和。未能找到文件“*\ShanXingTech.pfx”。  
解决方法：  
右键项目——属性——签名，去掉勾选 “为程序集签名(A)”，或者选择（或创建）你自己的强名称密钥文件。  


