# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
import random # 引入随机库，用于模拟人类行为和延迟

# --------------------------------------------------------

# 全局变量和路径
path=sys.path[0]+r'/1.txt'
num1 = 0

def gettoken(refresh_token):
    """
    使用 Refresh Token 换取新的 Access Token，并更新本地的 Refresh Token。
    """
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':id,
          'client_secret':secret,
          'redirect_uri':'http://localhost:53682/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    
    # 检查是否成功获取 Token
    if 'refresh_token' in jsontxt and 'access_token' in jsontxt:
        refresh_token = jsontxt['refresh_token']
        access_token = jsontxt['access_token']
        # 将新的 Refresh Token 写回文件，用于下次运行
        with open(path, 'w+') as f:
            f.write(refresh_token)
        return access_token
    else:
        # 如果 Token 刷新失败，打印错误信息并退出
        print("!!! 错误：Token 刷新失败。请检查 Refresh Token 或 Client Secret 是否正确。")
        print(jsontxt)
        sys.exit(1)


def main():
    """
    主函数：读取 Token -> 获取 Access Token -> 随机调用 API
    """
    try:
        fo = open(path, "r+")
        refresh_token = fo.read()
        fo.close()
    except FileNotFoundError:
        print("!!! 错误：未找到 1.txt 文件。请确保文件存在且包含有效的 Refresh Token。")
        return
        
    global num1
    localtime = time.asctime( time.localtime(time.time()) )
    
    # 1. 获取 Access Token
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    
    # --- A. 定义所有 GET 请求的 URL 列表 (读操作) ---
    api_urls = [
        r'https://graph.microsoft.com/v1.0/me/drive/root/children', # 文件列表
        r'https://graph.microsoft.com/v1.0/me/drive/recent', # 最近文件 (模拟浏览)
        r'https://graph.microsoft.com/v1.0/users', # 列出用户 (目录活动)
        r'https://graph.microsoft.com/v1.0/me/messages', # 邮件列表
        r'https://graph.microsoft.com/v1.0/me/mailFolders', # 邮件文件夹
        r'https://graph.microsoft.com/v1.0/me/calendar/events', # 日历事件
        r'https://graph.microsoft.com/v1.0/me/todo/lists', # 任务列表
        r'https://graph.microsoft.com/v1.0/me/onenote/notebooks', # OneNote 笔记本
        r'https://graph.microsoft.com/v1.0/me/contacts', # 联系人
        r'https://graph.microsoft.com/v1.0/me/memberOf', # 隶属于组
        r'https://graph.microsoft.com/v1.0/me/insights/trending', # 见解/趋势
        r'https://api.powerbi.com/v1.0/myorg/apps', # PowerBI
    ]
    
    # 2. 随机打乱 API 调用的顺序，模拟非固定模式
    random.shuffle(api_urls)
    
    try:
        # --- 3. 循环执行随机顺序的读取操作，并引入随机延迟 ---
        print(f"\n--- 本轮开始 ({localtime})，共 {len(api_urls)} 个读取 API ---")
        api_index = 0
        for url in api_urls:
            api_index += 1
            response = req.get(url, headers=headers)
            
            if response.status_code in [200, 201]:
                num1+=1
                print(f'{api_index:02d}. [GET] {url.split("/v1.0/")[-1]:<25} -> 成功 {num1} 次')
            else:
                print(f'{api_index:02d}. [GET] {url.split("/v1.0/")[-1]:<25} -> 失败 (状态码: {response.status_code})')
            
            # 引入随机停顿 (1到4秒) 模拟人查看页面的时间
            time.sleep(random.uniform(1, 4)) 

        # --- 4. 随机化写入操作的内容 (创建/删除草稿邮件) ---
        
        # 使用随机主题，避免每次写入的内容都完全一致
        current_time_str = time.strftime('%Y%m%d%H%M%S')
        draft_subject = f"Auto Draft KEEPALIVE {current_time_str}_{random.randint(100, 999)}"
        draft_data = {
            "subject": draft_subject, 
            "body": {"contentType": "Text", "content": "Checking mail write permissions by action."},
            "toRecipients": [{"emailAddress": {"address": "user@example.com"}}] 
        }

        # POST: 创建草稿邮件 (Mail.ReadWrite 权限)
        html_post = req.post(r'https://graph.microsoft.com/v1.0/me/messages', headers=headers, data=json.dumps(draft_data))
        if html_post.status_code == 201: # 201 表示创建成功
            num1+=1
            print(f'{api_index + 1:02d}. [POST] me/messages             -> 成功 {num1} 次 (创建草稿)')
            time.sleep(random.uniform(1, 3)) # 模拟人创建后的一小段时间间隔
            
            # DELETE: 清理草稿邮件 (Mail.ReadWrite 权限)
            try:
                message_id = json.loads(html_post.text)['id']
                # DELETE 成功通常返回 204 No Content
                if req.delete(f'https://graph.microsoft.com/v1.0/me/messages/{message_id}', headers=headers).status_code == 204:
                    num1+=1
                    print(f'{api_index + 2:02d}. [DELETE] me/messages/{message_id[:8]}... -> 成功 {num1} 次 (删除草稿)')
            except Exception as e:
                print("清理草稿邮件失败.")
        else:
             print(f'{api_index + 1:02d}. [POST] me/messages             -> 失败 (状态码: {html_post.status_code})')

        print('此次运行结束时间为 :', localtime)
        
    except Exception as e:
        print(f"主程序运行出错: {e}")
        pass

# --- 5. 执行循环，并引入更长的随机停顿 ---
for i in range(3):
    main()
    # 每次主函数运行结束后，引入10到30秒的随机停顿
    if i < 2:
        sleep_time = random.uniform(10, 30)
        print(f"\n--- 暂停 {sleep_time:.2f} 秒后进行下一轮循环... ---")
        time.sleep(sleep_time)
