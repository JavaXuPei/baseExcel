
#### 0.0.1
合并上下内容相同的单元格。

根据配置文件，设置标题大小，字体颜色。

根据传入参数，忽略指定列不需要合并。



```java
POST http://localhost:3000/index
Accept: application/json

{"rowTitleName": [{"sheetName":"Sheet1","rowTitleName": "成绩","isMergeCell": "N"}]}
```



![1.png](https://cdn.nlark.com/yuque/0/2021/png/533288/1626428082081-ce078d36-5750-4a0e-8c76-e0783ad2e207.png)

![2.png](https://cdn.nlark.com/yuque/0/2021/png/533288/1626428084131-3b172116-8eb7-4803-a586-03e7e7678a60.png)

