# Outlook-QA：提取 Outlook 邮件内容并生成 QA 语料库

本项目通过 Microsoft Graph API 和 OAuth 获取 Outlook 邮箱的邮件内容，并将其处理为适用于模型训练的 QA 语料库。

---

## 项目概览

- **目标**：从 Outlook 邮箱提取邮件，清洗数据，生成 QA 数据集。
- **方法**：使用 Microsoft Graph API 和 OAuth 协议访问邮箱。
- **输出**：以 CSV 格式保存的结构化 QA 语料库。

---

## 前置条件

- 拥有 Microsoft 账号及 Outlook 邮箱权限。
- 在 Azure 门户注册应用程序以获取 OAuth token。
- 配置 Python 环境，安装所需库（`msgraph-sdk`、`requests` 等）。

---

## 操作步骤

### 1. 获取 Token

Microsoft 已停止支持直接使用账号密码进行 IMAP 协议连接，需通过 OAuth 获取 token。

#### 方法 1：委托式获取 Token（手动）
1. 访问 [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)。
2. 登录并生成具有 `Mail.ReadWrite` 权限的 token。

#### 方法 2：应用程序获取 Token（程序化）
1. 在 Azure 门户注册一个应用程序。
2. 为应用授予 `Mail.ReadWrite` 权限。
3. 使用客户端凭据流获取 token。

---

### 2. 邮件操作

#### 方法 1：Microsoft Graph API
- **文档**：参考 [Microsoft Graph API 文档](https://docs.microsoft.com/en-us/graph/api/resources/mail-api-overview)。
- **测试**：通过 Graph Explorer 测试 API。
- **示例**：使用 `/me/messages` 端点获取邮件。

#### 方法 2：Microsoft Graph SDK
- 使用 Python SDK（`msgraph-sdk`）。
- 示例代码：参考 [msgraph-training-python](https://github.com/microsoftgraph/msgraph-training-python)。

---

### 3. 数据处理流程

#### 步骤 1：获取并解析邮件
- 使用 Graph API 获取原始邮件数据。
- 将响应数据转换为 JSON 格式。

#### 步骤 2：数据过滤
- **过滤标准**：
  1. 排除单条会话邮件（无回复）。
  2. 删除 HTML 中的无关元素（表格、广告插图等）。
  3. 排除会话数 ≥ 2 且最后一条为转发的邮件。
  4. 处理并合并类型为转发但不显示原始内容的邮件。

#### 步骤 3：转换为对话格式
- 将邮件会话转换为 JSON 对话格式。
- 使用 LLM 批量提取 QA 对。
- 保存为 CSV 文件，包含列：`Question`、`Answer`。

#### 步骤 4：去重
- 对 QA 对进行哈希处理，移除重复项。

---

## 输出格式

- **文件**：`qa_corpus.csv`
- **结构**：
  ```csv
  Question,Answer
  "会议时间是什么时候？","明天中午 2 点。"
  