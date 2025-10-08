import os
import json
import gspread
import logging
import datetime
from flask import Flask, request, jsonify, Response, escape
from flask_cors import CORS 
from oauth2client.service_account import ServiceAccountCredentials

# 設定日誌等級，方便在 Render 上除錯
logging.basicConfig(level=logging.INFO)

# ----------------------------------------------------
# 設定 Google Sheets 參數
spreadsheet_id = "1IHyA7aRxGJekm31KIbuORpg4-dVY8XTOEbU6p8vK3y4"
WORKSHEET_NAME = "設備報修" 

# Google Sheets API 範圍
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# 全域變數用於儲存 gspread client 和工作表
client = None
sheet = None

def initialize_gspread():
    """初始化 Google Sheets 連線。"""
    global client, sheet
    
    if client:
        return True 

    try:
        creds_json = os.environ.get('SERVICE_ACCOUNT_CREDENTIALS')
        if not creds_json:
            logging.error("致命錯誤：找不到 SERVICE_ACCOUNT_CREDENTIALS 環境變數。")
            return False

        # 嘗試解析 JSON 憑證
        creds_dict = json.loads(creds_json)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(
            creds_dict,
            scope
        )
        client = gspread.authorize(creds)
        
        # 嘗試打開試算表並選取工作表
        sheet = client.open_by_key(spreadsheet_id).worksheet(WORKSHEET_NAME)
        logging.info(f"成功連線到 Google Sheets。工作表名稱: {WORKSHEET_NAME}")
        return True

    except Exception as e:
        logging.error(f"連線到 Google Sheets 或打開工作表時發生錯誤: {e}")
        return False

# ----------------------------------------------------
# Flask 應用程式設定
app = Flask(__name__)
# 啟用 CORS，允許所有來源的網頁呼叫您的 API
CORS(app) 

# 在應用程式第一次請求前先初始化 gspread
with app.app_context():
    initialize_gspread()

# ----------------------------------------------------
# 路由定義

# 1. 根路由：用於顯示 HTML 報修表單 (您的 index.html 內容)
@app.route('/')
def home():
    """
    回傳報修表單的 HTML 內容。
    """
    # HTML 內容開始 (這部分是您的報修表單)
    html_content = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>設備報修系統</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {
            font-family: 'Inter', sans-serif;
            background-color: #f4f7f9;
        }
    </style>
</head>
<body class="min-h-screen flex items-center justify-center p-4">

    <div class="w-full max-w-lg bg-white p-8 md:p-10 rounded-xl shadow-2xl">
        
        <div class="text-center mb-8">
            <svg xmlns="http://www.w3.org/2000/svg" class="h-10 w-10 text-indigo-600 mx-auto mb-3" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                <path stroke-linecap="round" stroke-linejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37a1.724 1.724 0 002.572-1.065z" />
                <path stroke-linecap="round" stroke-linejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
            </svg>
            <h1 class="text-3xl font-bold text-gray-900">設備報修單</h1>
            <p class="text-gray-500 mt-1">請填寫詳細資訊，以便我們快速處理。</p>
        </div>

        <div id="message-box" class="hidden mb-6 p-3 text-center rounded-lg font-medium transition-all duration-300"></div>

        <form id="repairForm" class="space-y-6">
            
            <div>
                <label for="reporter_name_input" class="block text-sm font-medium text-gray-700 mb-1">報修人姓名 (必填)</label>
                <input type="text" id="reporter_name_input" required class="w-full px-4 py-2 border border-gray-300 rounded-lg shadow-sm focus:ring-indigo-500 focus:border-indigo-500">
            </div>

            <div>
                <label for="location_input" class="block text-sm font-medium text-gray-700 mb-1">設備位置 / 教室名稱 (必填)</label>
                <input type="text" id="location_input" required class="w-full px-4 py-2 border border-gray-300 rounded-lg shadow-sm focus:ring-indigo-500 focus:border-indigo-500">
            </div>

            <div>
                <label for="problem_input" class="block text-sm font-medium text-gray-700 mb-1">問題詳細描述 (必填)</label>
                <textarea id="problem_input" rows="4" required class="w-full px-4 py-2 border border-gray-300 rounded-lg shadow-sm focus:ring-indigo-500 focus:border-indigo-500 resize-none"></textarea>
            </div>
            
            <div>
                <label for="teacher_select" class="block text-sm font-medium text-gray-700 mb-1">協辦老師 (選填)</label>
                <select id="teacher_select" class="w-full px-4 py-2 border border-gray-300 rounded-lg shadow-sm focus:ring-indigo-500 focus:border-indigo-500 bg-white">
                    <option value="無指定">-- 選擇協辦老師 (無指定) --</option>
                    <option value="詹老師">詹老師</option>
                    <option value="佘老師">佘老師</option>
                    <option value="陳老師">陳老師</option>
                </select>
            </div>

            <div>
                <button type="submit" id="submit-button" class="w-full flex justify-center py-2 px-4 border border-transparent rounded-lg shadow-md text-sm font-medium text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 transition duration-150 ease-in-out">
                    送出報修單
                </button>
            </div>
            
            <div class="text-center pt-4 border-t border-gray-100 mt-6">
                <a href="/tasks" class="text-sm font-medium text-indigo-600 hover:text-indigo-500">
                    前往學生報修任務清單 →
                </a>
            </div>
        </form>
        </div>

    <script>
        // 設定您的 API 網址，這在 Render 部署時會自動指向正確的主機
        const API_URL_BASE = window.location.origin; 
        const API_URL = API_URL_BASE + "/submit_report";

        const form = document.getElementById('repairForm');
        const submitButton = document.getElementById('submit-button');
        const messageBox = document.getElementById('message-box');

        // 顯示訊息函式（取代 alert）
        function showMessage(message, isSuccess) {
            messageBox.textContent = message;
            messageBox.classList.remove('hidden', 'bg-red-100', 'text-red-800', 'bg-green-100', 'text-green-800');
            
            if (isSuccess) {
                messageBox.classList.add('bg-green-100', 'text-green-800');
            } else {
                messageBox.classList.add('bg-red-100', 'text-red-800');
            }
            // 5 秒後隱藏訊息
            setTimeout(() => {
                messageBox.classList.add('hidden');
            }, 5000);
        }

        form.addEventListener('submit', async function(event) {
            // 阻止表單的預設提交行為
            event.preventDefault();

            // 鎖定按鈕並顯示載入狀態
            submitButton.disabled = true;
            submitButton.textContent = '正在送出...';
            submitButton.classList.add('opacity-50', 'cursor-not-allowed');

            try {
                // 1. 從表單中收集資料
                const reportData = {
                    "reporterName": document.getElementById('reporter_name_input').value,
                    "deviceLocation": document.getElementById('location_input').value,
                    "problemDescription": document.getElementById('problem_input').value,
                    "helperTeacher": document.getElementById('teacher_select').value
                };

                // 2. 發送 POST 請求
                const response = await fetch(API_URL, {
                    method: 'POST',
                    // 告知伺服器我們正在傳送 JSON 格式的資料
                    headers: {
                        'Content-Type': 'application/json' 
                    },
                    body: JSON.stringify(reportData) // 將 JavaScript 物件轉換為 JSON 字串
                });

                const result = await response.json();

                if (response.ok) {
                    // HTTP 狀態碼為 200-299，表示成功
                    showMessage(result.message, true);
                    // 清空表單
                    form.reset(); 
                } else {
                    // HTTP 狀態碼為 4xx 或 5xx，表示 API 發生錯誤
                    throw new Error(result.message || `API 錯誤：HTTP 狀態碼 ${response.status}`);
                }

            } catch (error) {
                // 處理網路錯誤或 API 返回的錯誤訊息
                console.error("提交失敗:", error);
                // 顯示錯誤訊息
                showMessage(`提交失敗: ${error.message}`, false);
            } finally {
                // 無論成功或失敗，都恢復按鈕狀態
                submitButton.disabled = false;
                submitButton.textContent = '送出報修單';
                submitButton.classList.remove('opacity-50', 'cursor-not-allowed');
            }
        });
    </script>
</body>
</html>
    """
    # HTML 內容結束
    return Response(html_content, mimetype='text/html')


# 2. API 路由：用於接收表單提交的資料 (寫入 A-F 列)
@app.route('/submit_report', methods=['POST'])
def submit_data_api():
    """
    接收來自網頁的 POST 請求，將 JSON 資料寫入 Google Sheets。
    """
    if not sheet:
        return jsonify({"status": "error", "message": "伺服器初始化失敗，無法連線至 Google Sheets。請檢查 log 訊息。"}), 500

    try:
        data = request.get_json()
    except Exception:
        logging.error("請求資料解析失敗：不是有效的 JSON 格式。")
        return jsonify({"status": "error", "message": "請求必須是 JSON 格式。請檢查網頁前端的 Content-Type。"}), 400
    
    try:
        # 從 JSON 資料中提取欄位
        reporterName = data.get('reporterName', 'N/A')
        deviceLocation = data.get('deviceLocation', 'N/A')
        problemDescription = data.get('problemDescription', 'N/A')
        helperTeacher = data.get('helperTeacher', '無指定') # E 列欄位

        if not all([reporterName != 'N/A', deviceLocation != 'N/A', problemDescription != 'N/A']):
            logging.error(f"缺少必要資料: {data}")
            return jsonify({"status": "error", "message": "缺少必要的報修資料（如報修人、地點或描述）。"}), 400

        # 獲取台灣時間
        utc_now = datetime.datetime.utcnow()
        taiwan_time = utc_now + datetime.timedelta(hours=8)
        timestamp = taiwan_time.strftime("%Y-%m-%d %H:%M:%S")

        # row 陣列中包含 6 個元素：時間戳記、姓名、位置、描述、協辦老師、狀態
        row = [
            timestamp, 
            str(reporterName),
            str(deviceLocation),
            str(problemDescription),
            str(helperTeacher), # 協辦老師 (E 列)
            "待處理" # 狀態 (F 列)
        ]
        
        # 將資料附加到工作表的最後一行
        sheet.append_row(row)
        
        logging.info(f"資料成功寫入：{row}")
        return jsonify({"status": "success", "message": "設備報修資料已成功送出！"}), 200
        
    except Exception as e:
        logging.error(f"寫入 Google Sheets 時發生錯誤: {e}")
        return jsonify({"status": "error", "message": f"提交失敗：{str(e)}，可能是 Sheets API 限制或連線問題。"}), 500


# 3. 學生任務頁面路由：回傳 HTML
@app.route('/tasks')
def student_tasks_page():
    """
    回傳學生任務清單的 HTML 頁面。
    """
    # HTML 內容開始 (這部分是學生任務清單)
    html_content = f"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>學生報修任務清單</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {{
            font-family: 'Inter', sans-serif;
            background-color: #f4f7f9;
        }}
        .task-card {{
            transition: transform 0.2s, box-shadow 0.2s;
        }}
        .task-card:hover {{
            transform: translateY(-2px);
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
        }}
    </style>
</head>
<body class="p-4">
    <div class="max-w-4xl mx-auto">
        
        <div class="text-center mb-8">
            <h1 class="text-4xl font-extrabold text-gray-900">學長姐任務清單</h1>
            <p class="text-gray-500 mt-2">點擊「回報已完成」按鈕回報進度，請勿擅自更改他人任務！</p>
            <div class="mt-4">
                 <a href="/" class="text-sm font-medium text-indigo-600 hover:text-indigo-500">
                    ← 回到報修表單
                </a>
            </div>
        </div>
        
        <div id="message-box" class="hidden mb-6 p-3 text-center rounded-lg font-medium transition-all duration-300"></div>

        <div id="tasks-container" class="space-y-4">
            <div class="text-center text-gray-500 p-8" id="loading-message">
                <svg class="animate-spin h-5 w-5 mr-3 inline-block text-indigo-500" viewBox="0 0 24 24">
                  <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                  <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
                正在載入任務...
            </div>
        </div>
    </div>

    <script>
        // 設定 API 基礎 URL
        const API_URL_BASE = window.location.origin; 
        const GET_TASKS_URL = API_URL_BASE + "/get_tasks";
        const UPDATE_STATUS_URL = API_URL_BASE + "/update_status";
        const tasksContainer = document.getElementById('tasks-container');
        const loadingMessage = document.getElementById('loading-message');
        const messageBox = document.getElementById('message-box');

        // 輔助函式：將字串中的 HTML 特殊字元轉義，防止 XSS 攻擊
        function escape(html) {{
            return html.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
        }}

        // 顯示訊息函式
        function showMessage(message, isSuccess) {{
            messageBox.textContent = message;
            messageBox.classList.remove('hidden', 'bg-red-100', 'text-red-800', 'bg-green-100', 'text-green-800');
            
            if (isSuccess) {{
                messageBox.classList.add('bg-green-100', 'text-green-800');
            }} else {{
                messageBox.classList.add('bg-red-100', 'text-red-800');
            }}
            // 5 秒後隱藏訊息
            setTimeout(() => {{
                messageBox.classList.add('hidden');
            }}, 5000);
        }}

        // 任務卡片生成函式
        function createTaskCard(task) {{
            // 狀態顏色
            let statusClass = '';
            let buttonText = '回報已完成';
            let isCompleted = false;

            if (task.status === '待處理' || task.status === '處理中') {{
                statusClass = 'bg-yellow-100 text-yellow-800';
            }} else if (task.status === '已完成') {{
                statusClass = 'bg-green-100 text-green-800';
                buttonText = '已結案 (已完成)';
                isCompleted = true;
            }} else {{
                statusClass = 'bg-gray-100 text-gray-800';
            }}

            const card = document.createElement('div');
            // *** 修正後的代碼 ***
            card.className = `task-card bg-white p-6 rounded-xl shadow-lg border-l-4 border-indigo-500 ${{isCompleted ? 'opacity-70' : ''}}`;
            
            card.innerHTML = `
                <div class="flex justify-between items-start mb-4">
                    <div>
                        <span class="px-3 py-1 text-sm font-semibold rounded-full $ {{statusClass}}">${escape(task.status)}</span>
                    </div>
                    <span class="text-sm text-gray-500">${escape(task.timestamp)}</span>
                </div>
                
                <h3 class="text-xl font-bold text-gray-900 mb-2">${escape(task.deviceLocation)} - ${escape(task.reporterName)} 報修</h3>
                
                <div class="space-y-2 text-gray-600 text-sm mb-4">
                    <p><strong>協辦老師:</strong> ${escape(task.helperTeacher)}</p>
                    <p class="text-gray-700"><strong>問題描述:</strong> ${escape(task.problemDescription)}</p>
                </div>
                
                <div class="pt-4 border-t border-gray-100">
                    <button 
                        data-row-index="${task.rowIndex}" 
                        data-current-status="${task.status}"
                        class="status-button w-full py-2 px-4 rounded-lg text-white font-medium shadow-md transition duration-150 ease-in-out $ {{isCompleted ? 'bg-gray-400 cursor-not-allowed' : 'bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500'}}"
                        $ {{isCompleted ? 'disabled' : ''}}
                    >
                        ${buttonText}
                    </button>
                </div>
            `;
            
            // 綁定按鈕事件
            if (!isCompleted) {{
                const button = card.querySelector('.status-button');
                button.addEventListener('click', handleStatusUpdate);
            }}

            return card;
        }

        // 處理狀態更新
        async function handleStatusUpdate(event) {{
            const button = event.currentTarget;
            const rowIndex = button.dataset.rowIndex;
            const newStatus = '已完成'; // 點擊按鈕一律更新為「已完成」
            
            button.disabled = true;
            button.textContent = '正在更新...';
            button.classList.add('opacity-50', 'cursor-not-allowed');

            try {{
                const response = await fetch(UPDATE_STATUS_URL, {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify({{ 
                        rowIndex: rowIndex, 
                        newStatus: newStatus 
                    }})
                }});

                const result = await response.json();

                if (response.ok) {{
                    showMessage(result.message, true);
                    // 成功後重新載入任務列表
                    await loadTasks(); 
                }} else {{
                    throw new Error(result.message || '更新失敗');
                }}

            }} catch (error) {{
                console.error("更新失敗:", error);
                showMessage(`更新失敗: ${error.message}`, false);
            }} 
        }


        // 載入任務清單
        async function loadTasks() {{
            tasksContainer.innerHTML = ''; // 清空舊列表
            loadingMessage.classList.remove('hidden');

            try {{
                const response = await fetch(GET_TASKS_URL);
                const result = await response.json();

                loadingMessage.classList.add('hidden');

                if (response.ok) {{
                    if (result.tasks.length === 0) {{
                        tasksContainer.innerHTML = '<p class="text-center text-gray-500 p-8">目前沒有任何報修任務。</p>';
                        return;
                    }}
                    
                    // 排序：未完成的排前面
                    const sortedTasks = result.tasks.sort((a, b) => {{
                        if (a.status === '已完成' && b.status !== '已完成') return 1;
                        if (a.status !== '已完成' && b.status === '已完成') return -1;
                        return 0; 
                    }});

                    sortedTasks.forEach(task => {{
                        tasksContainer.appendChild(createTaskCard(task));
                    }});

                }} else {{
                    throw new Error(result.message || '無法取得任務清單');
                }}

            }} catch (error) {{
                loadingMessage.classList.add('hidden');
                console.error("載入任務失敗:", error);
                tasksContainer.innerHTML = `<p class="text-center text-red-600 p-8">載入任務失敗：${error.message}</p>`;
            }}
        }}

        // 頁面載入時執行
        window.onload = loadTasks;
    </script>
</body>
</html>
    """
    # HTML 內容結束
    return Response(html_content, mimetype='text/html')

# 4. API 路由：用於讀取 Google Sheets 所有報修資料
@app.route('/get_tasks', methods=['GET'])
def get_tasks_api():
    """
    從 Google Sheets 讀取所有報修資料，並回傳 JSON 列表。
    """
    if not sheet:
        return jsonify({"status": "error", "message": "伺服器初始化失敗，無法連線至 Google Sheets。"}), 500

    try:
        # 讀取工作表中的所有資料
        all_data = sheet.get_all_values()
        
        # 假設第一行為標題，跳過
        if len(all_data) <= 1:
            return jsonify({"status": "success", "tasks": []})

        records = all_data[1:]
        
        tasks_list = []
        # 從第 2 列開始，所以 row_index 從 2 開始
        for i, row in enumerate(records, start=2): 
            # 確保 row 至少有 6 個元素 (時間, 姓名, 位置, 描述, 協辦老師, 狀態)
            if len(row) < 6:
                # 忽略不完整的報修記錄
                continue

            task = {
                "rowIndex": i, # 在 Sheets 中的實際列號，用於後續更新
                "timestamp": row[0],
                "reporterName": row[1],
                "deviceLocation": row[2],
                "problemDescription": row[3],
                "helperTeacher": row[4], # 協辦老師 (E 列, 索引 4)
                "status": row[5] # 狀態 (F 列, 索引 5)
            }
            tasks_list.append(task)
            
        return jsonify({"status": "success", "tasks": tasks_list}), 200
        
    except Exception as e:
        logging.error(f"讀取 Google Sheets 時發生錯誤: {e}")
        return jsonify({"status": "error", "message": f"讀取任務失敗：{str(e)}，請檢查 Sheets 權限。"}), 500

# 5. API 路由：用於更新報修記錄的狀態
@app.route('/update_status', methods=['POST'])
def update_status_api():
    """
    接收 POST 請求，根據列號 (rowIndex) 更新報修記錄的狀態 (F 列)。
    """
    if not sheet:
        return jsonify({"status": "error", "message": "伺服器初始化失敗，無法連線至 Google Sheets。"}), 500
    
    try:
        data = request.get_json()
        rowIndex = int(data.get('rowIndex'))
        newStatus = data.get('newStatus')
        
        if not rowIndex or not newStatus or rowIndex < 2:
            return jsonify({"status": "error", "message": "無效的請求資料：缺少列號或新狀態。"}), 400

        # 「狀態」欄位對應 Sheets 的 F 列，在 gspread 中列號 (col) 從 1 開始數，所以 F 列是 6
        STATUS_COLUMN_INDEX = 6 

        sheet.update_cell(rowIndex, STATUS_COLUMN_INDEX, newStatus)
        
        logging.info(f"成功更新第 {rowIndex} 列的狀態為: {newStatus}")
        return jsonify({"status": "success", "message": f"任務狀態已成功更新為「{newStatus}」！"}), 200
        
    except Exception as e:
        logging.error(f"更新 Google Sheets 時發生錯誤: {e}")
        return jsonify({"status": "error", "message": f"更新狀態失敗：{str(e)}。"}), 500

# ----------------------------------------------------
# 本地測試運行
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=os.environ.get('PORT', 5000))
