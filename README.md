# outlook-qa
Get all email content from Outlook and generate a QA corpus.

使用imap协议获取outlook邮箱邮件，处理为模型可用的语料库



## 说明
### 1. token获取
目前微软已取消使用账号密码配置的imap协议连接，改成了oauth
#### 1. 委托式获取token
https://developer.microsoft.com/en-us/graph/graph-explorer
#### 2. 应用程序获取token
创建应用程序，开放邮箱读写权限
### 2. 邮件操作
#### 1. api
api文档与测试见graph-explorer

#### 2. sdk
见msgraph-training-python

### 3. 数据清洗
（1）请求获取邮件api，读取原始数据，转换json

（2）过滤：

    a. 单条会话(即未回复的邮件)
    b. 删除邮件中html不相关的元素：删除表格元素、广告插图
    c. 会话数为2+末条邮件类型为转发
    d. 处理与合并类型为转发、不显示原始内容的邮件
（3） 转换为对话json，llm批量提取，写入csv

（4） 哈希去重qa
