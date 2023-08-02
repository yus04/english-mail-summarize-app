import os, logging, openai

import smtplib
from email import charset
from email.mime.text import MIMEText

import azure.functions as func
from azure.storage.queue import QueueClient, QueueMessage
from azure.core.paging import ItemPaged

# Queueサービスに接続するためのアカウント情報を設定
account_name = os.getenv("ACCOUNT_NAME")
account_key = os.getenv("ACCOUNT_KEY")
account_url = "https://" + account_name + ".queue.core.windows.net/"

openai.api_type = "azure"
openai.api_version = "2023-05-15"
openai.api_base = os.getenv('AOAI_BASE', None)
openai.api_key = os.getenv('AOAI_APIKEY', None)

# Outlook設定
mail_account = os.getenv("MAIL_ACCOUNT")
mail_password = os.getenv("MAIL_PASSWORD")

# メールの設定
mail_to = os.getenv("MAIL_TO")
mail_subject = 'English Mail Summarize Master'
charset.add_charset('utf-8', charset.SHORTEST, None, 'utf-8')

# デキューするQueueの名前を指定
queue_name = os.getenv("QUEUE_NAME_FOR_TIMER")

# デキューしたいメッセージの最大数を指定
max_messages = 4

# Queueサービスに接続
queue_client = QueueClient(
    account_url=account_url,
    queue_name=queue_name,
    credential=account_key
)

def main(timer: func.TimerRequest) -> None:
    logging.info('Python Timer trigger function processed a request.')

    # メッセージの取得
    messages = queue_client.receive_messages(max_messages=max_messages)
    logging.info(f"Peeked messages: {messages}")

    # chatGPTへの問い合わせ
    answers_dict = ask_chat_gpt(messages)
    logging.info(f"Answers dict: {answers_dict}")

    # メッセージのデキュー
    dequeue(messages)

    # 要約メッセージの作成
    summarized_answers_string = summarized_answers(answers_dict)

    # メール送信
    send_message(summarized_answers_string)

    # return func.HttpResponse(summarized_answers_string)

def ask_chat_gpt(messages: ItemPaged[QueueMessage]) -> dict:
    answers_dict = {}
    # ピークしたメッセージそれぞれに対して、chatGPTへの問い合わせを行う
    for message in messages:
        content = message.content.replace("\n", "")
        logging.info(f"Peeked message content: {content}")

        # chatGPTへの問い合わせ
        response = chat_gpt(message.content)
        response = chat_gpt_for_summarize(response)
        logging.info(f"Response message: {response}")

        answers_dict[content] = response

    return answers_dict

def chat_gpt(input_txt: str) -> str:
    response = openai.ChatCompletion.create(
        engine="gpt-35-turbo",
        messages = [
            {"role":"system","content":"英語の文章を次の条件で要約してください。\n- 日本語で回答\n- 50文字程度で回答\n- 1文で回答"},
            {"role":"user","content":"Local governments in Japan are embracing ChatGPT, the generative AI chatbot developed by US venture firm OpenAI.\
             Yokosuka, a city south of Tokyo, has become the first to implement the language model in all of its offices on an experimental basis."},
            {"role":"assistant","content":"横須賀市が先駆けでChatGPTを実験的に導入。"},
            {"role": "user", "content": input_txt},
        ],
        temperature=0.0,
        max_tokens=800,
        frequency_penalty=0,
        presence_penalty=0,
        stop=None)
    return response['choices'][0]['message']['content']

def chat_gpt_for_summarize(input_txt: str) -> str:
    response = openai.ChatCompletion.create(
        engine="gpt-35-turbo",
        messages = [
            {"role":"system","content":"文章を100文字程度で要約してください。"},
            {"role": "user", "content": input_txt},
        ],
        temperature=0.0,
        max_tokens=800,
        frequency_penalty=0,
        presence_penalty=0,
        stop=None)
    return response['choices'][0]['message']['content']

def dequeue(messages: ItemPaged[QueueMessage]) -> None:
    for message in messages:
        # メッセージの処理が終わったら、メッセージを削除
        queue_client.delete_message(message.id, message.pop_receipt)
        logging.info(f"Deleted message: {messages}")

def summarized_answers(answers_dict: dict) -> str:
    summarized_answers_string = ""
    for i, value in enumerate(answers_dict.values()):
        summarized_answers_string += f"{i+1}つ目のメールの要約結果\n"
        summarized_answers_string += str(value)
        summarized_answers_string += "\n\n"
    return summarized_answers_string

def send_outlook_mail(msg: MIMEText) -> None:
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(mail_account, mail_password)
    server.send_message(msg)

def make_mime(mail_to: str, mail_subject: str, body: str) -> MIMEText:
    msg = MIMEText(body, 'plain', "utf-8")
    msg['Subject'] = mail_subject
    msg['To'] = mail_to
    msg['From'] = mail_account
    return msg

def send_message(body: str) -> None:
    msg = make_mime(
        mail_to=mail_to,
        mail_subject=mail_subject,
        body=body)
    send_outlook_mail(msg)
