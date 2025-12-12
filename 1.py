# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
import random # 引入随机库，用于模拟人类行为

# ... (注册权限的注释不变) ...

path=sys.path[0]+r'/1.txt'
num1 = 0

def gettoken(refresh_token):
    # ... (gettoken 函数保持不变，因为它是获取Token的核心逻辑) ...
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
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token

def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    
    # --- 1. 定义所有 GET 请求的 URL 列表 (读操作) ---
    # 增加更多服务种类，提升多样性
    api_urls = [
        r'https://graph.microsoft.com/v1.0/me/drive/root',
        r'https://graph.microsoft.com/v1.0/me/drive/root/children', # 文件列表
        r'https://graph.microsoft.com/v1.0/me/drive/recent', # 最近文件 (更像人)
        r'https://graph.microsoft.com/v1.0/users', # 列出用户 (已修复空格)
        r'https://graph.microsoft.com/v1.0/me/messages', # 邮件列表
        r'https://graph.microsoft.com/v1.0/me/mailFolders', # 邮件文件夹
        r'https://graph.microsoft.com/v1.0/me/calendar/events', # 日历事件
        r'https://graph.microsoft.com/v1.0/me/todo/lists', # 任务列表
        r'https://graph.microsoft.com/v1.0/me/onenote/notebooks', # OneNote 笔记本
        r'https://graph.microsoft.com/v1.0/me/contacts', # 联系人
        r'https://graph.microsoft.com/v1.0/me/sites', # SharePoint 网站
        r'https://graph.microsoft.com/v1.0/me/memberOf', # 隶属于组
        r'https://graph.microsoft.com/v1.0/me/insights/trending', # 见解/趋势
        r'https://api.powerbi.com/v1.0/myorg/apps', # PowerBI
    ]
    
    try:
        # --- 2. 随机打乱 API 调用的顺序 ---
        random.shuffle(api_urls)
        
        # --- 3. 循环执行随机顺序的读取操作，并引入随机延迟 ---
        api_index = 0
        for url in api_urls:
            api_index += 1
            response = req.get(url, headers=headers)
            
            if response.status_code in [200, 201]:
                num1+=1
                # 简化打印信息，只保留索引和成功次数
                print(f'{api_index}.调用成功 {num1} 次') 
            
            # 引入随机停顿 (1到5秒) 模拟人查看页面的时间
            time.sleep(random.uniform(1, 5)) 

        # --- 4. 随机化写入操作的内容 (创建/删除草稿邮件) ---
        
        # 使用随机主题，避免每次写入的内容都完全一致
        draft_subject = f"Auto Draft by GitHub Action {time.strftime('%Y%m%d')}_{random.randint(1000, 9999)}"
        draft_data = {
            "subject": draft_subject, 
            "body": {"contentType": "Text", "content": "Checking mail write permissions."},
            "toRecipients": [{"emailAddress": {"address": "user@example.com"}}] 
        }

        # POST: 创建草稿邮件
        html_post = req.post(r'https://graph.microsoft.com/v1.0/me/messages', headers=headers, data=json.dumps(draft_data))
        if html_post.status_code == 201: # 201 表示创建成功
            num1+=1
            print(f'{api_index + 1}.调用成功 {num1} 次 (创建草稿邮件: POST)')
            time.sleep(random.uniform(1, 3)) # 模拟人创建后的一小段时间间隔
            
            # DELETE: 清理草稿邮件
            try:
                message_id = json.loads(html_post.text)['id']
                # DELETE 成功通常返回 204 No Content
                if req.delete(f'https://graph.microsoft.com/v1.0/me/messages/{message_id}', headers=headers).status_code == 204:
                    num1+=1
                    print(f'{api_index + 2}.调用成功 {num1} 次 (删除草稿邮件: DELETE)')
            except Exception as e:
                print("清理草稿邮件失败.")

        print('此次运行结束时间为 :', localtime)
        
    except Exception as e:
        print(f"主程序运行出错: {e}")
        pass

# --- 5. 循环之间引入更长的随机停顿 ---
for i in range(3):
    main()
    # 每次主函数运行结束后，引入10到30秒的随机停顿
    if i < 2:
        sleep_time = random.uniform(10, 30)
        print(f"暂停 {sleep_time:.2f} 秒后进行下一轮循环...")
        time.sleep(sleep_time)
