![header](./img/header.png)

# sAIten
Templates and scripts for creating and administering AI-assessable quizzes for educational use



# Requirement
* Google account
* Gemini AI API Key


# Example Workflow (Personal ver.)
```mermaid
sequenceDiagram
    autonumber
    participant Google as Google
    participant Repo as sAIten Repository
    actor Author as Question Author
    participant Main as Authoring Sheet
    participant Score as Summary Sheet
    participant Form as Assignment Form
    actor Student as Test taker
    participant Gemini as Gemini AI

    Note over Google, Main: [Deployment Phase]
    Author->>Google: Request for an API key
    Google->>Author: Generate API key
    Author->>Repo: Download
    Repo-->>Main: Stored in Google Drive
    Author->>Main: Initialization and deployment

    Note over Author, Form: [Create questions]
    Author->>Main: 設問と評価基準の記入
    Author-->>Main: 設定の確認・変更
    Main->>Score: 集計シートの生成
    Main->>Form: フォームの生成

    Note over Author, Form: 【確認フェーズ】
    Author->>Form: 問題の確認、図表の追加
    Author->>Score: 設定の確認（回答者、期限の設定など）

    Note over Author,Gemini : 【実行フェーズ】
    Author->>Student: フォームURLの通知
    Student->>Form: 解答の送信
    Form->>Score: 回答を収集
    Score->>Gemini: 回答と評価基準を送信
    Gemini->>Score: 採点結果（評点・コメント）を返却
    Score->>Student: 採点結果の自動送信

    Student->>Author: 採点結果についての質問
    Author->>Student: 質問への回答...

```
  

# Usage
1. Copy the [spread sheet]() to your GoogleDrive.

![demo video]()

The sheet is 
\* For clear managing, I recommend create a new directory for saving quiz forms and score sheets of them.

1. Open the copied sheet and enter the information of your quiz and .
![thmubnail_0x]()


1. Select "" > "" menu from the menubar.
![thmubnail_0x]()

Now, two files 

1. Check the created form.

![demo video]()


1. Open the created scoring sheet and select "" > "" menu to authorize your account.
During this step, you need to enter your Gemini AI API key for automatic scoring by AI.

1. Next, select "" > "" menu to authoriza your account for automatic sending scoring results to students.

\* The authorization steps for AI usage and sending emails are separated because  
Please run both the authorization steps.

1. (Optional) Edit the setting.


# Sample



