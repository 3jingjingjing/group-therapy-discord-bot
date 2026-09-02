import discord
from discord.ext import commands
import random
import os
import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv
import asyncio
from datetime import datetime

load_dotenv()  # 載入 .env 檔案中的環境變數（PM2 的 env 會覆蓋，但這確保前台運行也 ok）

# 定義角色映射和 Prompt（根據你的 5 個角色自訂 prompt）
ROLES = {
    'role1': {'token': os.getenv('DISCORD_BOT_TOKEN_ROLE1'), 'prompt': '你是一個友好、熱情的助手，總是用正面語言回應。', 'bot_name': '3jingㄉ論文孕育實驗室#8326'},
    'role3': {'token': os.getenv('DISCORD_BOT_TOKEN_ROLE3'), 'prompt': '你是一個嚴肅、專業的顧問，專注於事實和建議。', 'bot_name': '角色三_3jing論文孕育實驗室#4873'},
    'role4': {'token': os.getenv('DISCORD_BOT_TOKEN_ROLE4'), 'prompt': '你是一個幽默、搞笑的角色，總是用笑話回應。', 'bot_name': '角色四_3jing論文孕育實驗室#1310'},
    'role5': {'token': os.getenv('DISCORD_BOT_TOKEN_ROLE5'), 'prompt': '你是一個創意、想像力的藝術家，總是用詩意語言。', 'bot_name': '角色五_3jing論文孕育實驗室#6613'},
    'role_care': {'token': os.getenv('DISCORD_BOT_TOKEN_ROLE_CARE'), 'prompt': '你是一個關懷、支持性的角色，總是用同理心回應。', 'bot_name': '關心角色_3jing論文孕育實驗室#5514'}
}

# LLM 客戶端
client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

# Excel 檔案路徑
HISTORY_EXCEL = 'conversation_history.xlsx'

# 追蹤每個使用者的 last_mentioned_role（用 dict 記憶）
last_mentioned_roles = {}  # key: user_id, value: role (如果 @ 了特定 Bot，下次用這個 role)

# 多 Client 管理
clients = {}
for role, data in ROLES.items():
    intents = discord.Intents.default()
    intents.message_content = True
    intents.messages = True
    intents.members = True  # 需要 members intent 來檢查 mentions
    bot = commands.Bot(command_prefix='!', intents=intents)
    bot.role = role  # 附加角色資訊
    clients[role] = bot

# 登入所有 Bot（使用 gather 並行運行）
async def login_all_bots():
    tasks = []
    for role, bot in clients.items():
        token = ROLES[role]['token']
        if not token:
            print(f"Warning: No token for {role}")
            continue
        tasks.append(bot.start(token))
    await asyncio.gather(*tasks)

# 更新歷史：新增一筆記錄到使用者的 sheet
def append_to_history(user_id, sender, message):
    try:
        if not os.path.exists(HISTORY_EXCEL):
            # 如果檔案不存在，建立空檔案
            with pd.ExcelWriter(HISTORY_EXCEL, engine='openpyxl') as writer:
                pd.DataFrame(columns=['user_id', 'timestamp', 'sender', 'message']).to_excel(writer, sheet_name=user_id, index=False)
        else:
            excel_file = pd.ExcelFile(HISTORY_EXCEL)
            if user_id not in excel_file.sheet_names:
                # 如果 sheet 不存在，建立空 sheet
                with pd.ExcelWriter(HISTORY_EXCEL, mode='a', engine='openpyxl') as writer:
                    pd.DataFrame(columns=['user_id', 'timestamp', 'sender', 'message']).to_excel(writer, sheet_name=user_id, index=False)

        # 讀取現有 sheet
        df = pd.read_excel(HISTORY_EXCEL, sheet_name=user_id, engine='openpyxl')
        
        new_row = {
            'user_id': user_id,
            'timestamp': datetime.now().isoformat(),
            'sender': sender,
            'message': message
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        
        # 寫回
        with pd.ExcelWriter(HISTORY_EXCEL, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=user_id, index=False)
    except Exception as e:
        print(f"Error appending to history for user {user_id}: {e}")

# 獲取特定 user's 總 user message_count
def get_message_count(user_id):
    if not os.path.exists(HISTORY_EXCEL) or user_id not in pd.ExcelFile(HISTORY_EXCEL).sheet_names:
        return 0
    df = pd.read_excel(HISTORY_EXCEL, sheet_name=user_id, engine='openpyxl')
    user_df = df[df['sender'] == 'user']
    return len(user_df)

# 獲取特定 user's 最新 summary
def get_summary(user_id):
    if not os.path.exists(HISTORY_EXCEL) or user_id not in pd.ExcelFile(HISTORY_EXCEL).sheet_names:
        return ''
    df = pd.read_excel(HISTORY_EXCEL, sheet_name=user_id, engine='openpyxl')
    summary_rows = df[df['sender'] == 'summary']
    if not summary_rows.empty:
        return summary_rows.iloc[-1]['message']  # 最新 summary
    return ''

# 生成或更新 summary
def update_summary(user_id):
    if not os.path.exists(HISTORY_EXCEL) or user_id not in pd.ExcelFile(HISTORY_EXCEL).sheet_names:
        return
    df = pd.read_excel(HISTORY_EXCEL, sheet_name=user_id, engine='openpyxl')
    full_history = '\n'.join(df.apply(lambda row: f"{row['sender']}: {row['message']}", axis=1))
    summary_prompt = f"總結以下對話歷史為簡短大綱:\n{full_history}"
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "system", "content": summary_prompt}]
        )
        new_summary = response.choices[0].message.content
        append_to_history(user_id, 'summary', new_summary)
    except Exception as e:
        print(f"Error generating summary for user {user_id}: {e}")
        append_to_history(user_id, 'summary', "Summary generation failed.")

# 獲取特定 role 的歷史（user messages + 该 role 的 responses）
def get_role_history(user_id, role):
    if not os.path.exists(HISTORY_EXCEL) or user_id not in pd.ExcelFile(HISTORY_EXCEL).sheet_names:
        return ''
    df = pd.read_excel(HISTORY_EXCEL, sheet_name=user_id, engine='openpyxl')
    role_df = df[(df['sender'] == 'user') | (df['sender'] == role)]
    return '\n'.join(role_df.apply(lambda row: f"{row['sender']}: {row['message']}", axis=1))

# 生成回應的 LLM 函數（加入 role-specific history）
def generate_response(role, user_message, summary, role_history):
    system_prompt = ROLES[role]['prompt'] + f"\n對話大綱: {summary}\n你的先前對話紀錄: {role_history}\n回應使用者訊息，保持個性一致。"
    try:
        response = client.chat.completions.create(
            model="gpt-4o",  # 或其他模型
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message}
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error generating response for {role}: {e}")
        return "Sorry, I couldn't generate a response."

# 管理 LLM 決定角色（移除 JSON，改成簡單文字輸出）
def select_roles(user_message, summary, num_roles):
    manage_prompt = f"根據使用者訊息: '{user_message}' 和對話大綱: '{summary}'，決定適合的 {num_roles} 個角色。角色選項: {list(ROLES.keys())}。輸出角色名稱用空格分隔，如 role1 role3。只輸出角色名稱，絕對無額外文字。"
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "system", "content": manage_prompt}]
        )
        content = response.choices[0].message.content.strip()
        selected = content.split()  # 用空格分隔成 list
        return selected[:num_roles]  # 確保數量
    except Exception as e:
        print(f"Error selecting roles: {e}")
        return random.sample(list(ROLES.keys()), num_roles)  # fallback

# 主訊息處理（用 role_care 作為主接收器）
main_bot = clients['role_care']

@main_bot.event
async def on_ready():
    print(f"Main bot {main_bot.user} is ready!")

@main_bot.event
async def on_message(message):
    if message.author.bot:
        return
    user_id = str(message.author.id)  # 以 user_id 為 key
    user_message = message.content

    # 檢查是否有 @ 特定 Bot
    mentioned_role = None
    for mentioned_user in message.mentions:
        mentioned_name = f"{mentioned_user.name}#{mentioned_user.discriminator}"
        for role, data in ROLES.items():
            if data['bot_name'] == mentioned_name:
                mentioned_role = role
                last_mentioned_roles[user_id] = role  # 記錄為該 user 的 last_mentioned_role
                break
        if mentioned_role:
            break

    # 更新歷史：新增 user message
    append_to_history(user_id, 'user', user_message)

    # 檢查是否需要生成 summary
    message_count = get_message_count(user_id)
    if message_count % 10 == 0:
        update_summary(user_id)

    # 獲取最新 summary
    summary = get_summary(user_id)

    # 步驟1: 亂數決定回覆數量 (1 or 2)
    num_responses = random.choice([1, 2])

    # 步驟2: 管理 LLM 決定角色
    selected_roles = select_roles(user_message, summary, num_responses)

    # 如果有 last_mentioned_role，強制用它作為下一個回應（並清空，避免持續）
    if user_id in last_mentioned_roles:
        selected_roles = [last_mentioned_roles[user_id]]  # 只用這個 role 回應
        del last_mentioned_roles[user_id]  # 清空，下次恢復正常

    # 步驟3: 為每個角色生成回應並發送，並記錄到歷史
    for role in selected_roles:
        role_history = get_role_history(user_id, role)
        response_text = generate_response(role, user_message, summary, role_history)
        bot = clients.get(role)
        if bot and bot.is_ready():
            try:
                channel = await bot.fetch_channel(message.channel.id)
                await channel.send(response_text)
                # 新增 bot response 到歷史
                append_to_history(user_id, role, response_text)
            except Exception as e:
                print(f"Error sending message from {role}: {e}")
        else:
            print(f"Bot for {role} not ready or not found.")
        await asyncio.sleep(1)  # 限速，避免 rate limit

    await main_bot.process_commands(message)

# 運行所有 Bot
if __name__ == "__main__":
    asyncio.run(login_all_bots())