import email
from email import policy
from email.parser import BytesParser
from email.header import decode_header
import os



def parse_eml(eml_path):
    with open(eml_path, 'rb') as f:  # 以二进制模式打开文件
        # 使用 BytesParser 解析邮件内容
        msg = BytesParser(policy=policy.default).parse(f)

    # 提取基础信息 -------------------------------------------------
    subject, encoding = decode_header(msg['Subject'])[0]
    if encoding:
        subject = subject.decode(encoding)

    from_, _ = decode_header(msg.get('From'))[0]
    date = msg.get('Date')

    print(f"主题: {subject}")
    print(f"发件人: {from_}")
    print(f"日期: {date}")

    # 提取正文内容 -------------------------------------------------
    body = ""
    if msg.is_multipart():
        # 遍历邮件各部分（正文、附件等）
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            # 跳过附件部分
            if "attachment" in content_disposition:
                continue

            # 提取文本或HTML正文
            if content_type in ["text/plain", "text/html"]:
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset()
                if charset:
                    body += payload.decode(charset, errors="ignore")
                else:
                    body += payload.decode("utf-8", errors="ignore")
    else:
        # 非多部分邮件的直接读取
        payload = msg.get_payload(decode=True)
        charset = msg.get_content_charset()
        body = payload.decode(charset or "utf-8", errors="ignore")

    print("正文内容:")
    print(body[:500] + "...")  # 只打印前500字符避免控制台卡顿

    # 提取附件 -----------------------------------------------------
    for part in msg.iter_attachments():
        filename = part.get_filename()
        if filename:
            filename, encoding = decode_header(filename)[0]
            if encoding:
                filename = filename.decode(encoding)
            # 保存附件到当前目录
            with open(filename, 'wb') as f:
                f.write(part.get_payload(decode=True))
            print(f"附件已保存: {filename}")

    return {
        "subject": subject,
        "from": from_,
        "date": date,
        "body": body,
        "attachments": [part.get_filename() for part in msg.iter_attachments()]
    }


# 使用示例
eml_file = "/Users/txmmy/my-python-projects/outlook-qa/已添加 Microsoft 帐户安全信息.eml"
result = parse_eml(eml_file)
