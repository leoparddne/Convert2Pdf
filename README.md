# Convert2PDF

#### 介绍
使用LibreOffice 转换excel/ppt/word到PDF
支持xls,xlsx,ppt,pptx,doc,docx

核心为调用soffice,所以需要提前安装libreoffice  
下文提到的soffice即libreoffice的可执行程序，可以在安装目录获取(不同版本可能文件名不同,使用主程序即可)

```
soffice --invisible --convert-to pdf  "PDF文件所在位置" --outdir  "输出文件夹"
```


#### 示例
```
    ConverterHelper convertHelper = new ConverterHelper();

    string sofficePath = @"D:\Program Files\LibreOffice\program\soffice.exe";
    convertHelper.Init(sofficePath);

    var result = convertHelper.ConverterToPDF(@"C:\Users\ivesBao\Desktop\test.docx", @"C:\Users\ivesBao\Desktop\testpdf");
    Console.WriteLine(result);
```

#### 参与贡献

1.  Fork 本仓库
2.  新建 Feat_xxx 分支
3.  提交代码
4.  新建 Pull Request


#### 特技

1.  使用 Readme\_XXX.md 来支持不同的语言，例如 Readme\_en.md, Readme\_zh.md
2.  Gitee 官方博客 [blog.gitee.com](https://blog.gitee.com)
3.  你可以 [https://gitee.com/explore](https://gitee.com/explore) 这个地址来了解 Gitee 上的优秀开源项目
4.  [GVP](https://gitee.com/gvp) 全称是 Gitee 最有价值开源项目，是综合评定出的优秀开源项目
5.  Gitee 官方提供的使用手册 [https://gitee.com/help](https://gitee.com/help)
6.  Gitee 封面人物是一档用来展示 Gitee 会员风采的栏目 [https://gitee.com/gitee-stars/](https://gitee.com/gitee-stars/)
