Office.onReady(() => {
    // Office SDKの準備完了
});

// Entra ID (Azure AD) 認証設定
const msalConfig = {
    auth: {
        // ★ここをAzure Portalで取得したクライアントIDに書き換える★
        clientId: "f6ff460c-788d-4dd2-998d-3369c941ae11", 
        // ★ここをAzure Portalで取得したテナントIDに書き換える★
        authority: "https://login.microsoftonline.com/1388e299-7a0d-4bf5-9f1c-49f365d23f7c", 
    }
};

const msalApp = new msal.PublicClientApplication(msalConfig);

// メイン関数：ボタンクリック時に呼び出される
async function moveAndCreateRule(event) {
    const item = Office.context.mailbox.item;
    const senderEmail = item.from.emailAddress;
    // ★ここを移動先のフォルダIDに書き換える（例: "Archive", "Inbox", もしくは固有のフォルダID文字列）★
    const targetFolderId = "YOUR_TARGET_FOLDER_ID"; 

    try {
        // 1. Graph API アクセストークンの取得
        const token = await getGraphToken();

        // 2. 現在のメッセージを移動
        await moveMessage(item.itemId, targetFolderId, token);

        // 3. 次回以降の自動移動ルールを作成
        await createInboxRule(senderEmail, targetFolderId, token);

        // 成功通知の表示
        Office.context.mailbox.item.notificationMessages.addAsync("success", {
            type: "informationalMessage",
            message: `メールを移動し、${senderEmail} の自動移動ルールを作成しました。`,
            icon: "icon16",
            persistent: false
        });

    } catch (error) {
        console.error("Error executing action:", error);
        Office.context.mailbox.item.notificationMessages.addAsync("error", {
            type: "errorMessage",
            message: "処理中にエラーが発生しました。"
        });
    } finally {
        event.completed(); // Officeへ処理完了を通知
    }
}

// Graph API トークン取得関数
async function getGraphToken() {
    const request = {
        scopes: ["Mail.ReadWrite", "MailboxSettings.ReadWrite"]
    };
    try {
        const response = await msalApp.acquireTokenPopup(request);
        return response.accessToken;
    } catch (error) {
        console.error("Token acquisition failed:", error);
        throw error;
    }
}

// メール移動処理
async function moveMessage(messageId, targetFolderId, token) {
    const restId = Office.context.mailbox.convertToRestId(messageId, Office.MailboxEnums.RestVersion.v2_0);
    
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}/move`, {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ "destinationId": targetFolderId })
    });

    if (!response.ok) {
        throw new Error(`Failed to move message: ${response.statusText}`);
    }
}

// 受信トレイ ルール作成処理
async function createInboxRule(senderEmail, targetFolderId, token) {
    const rulePayload = {
        "displayName": `自動移動: ${senderEmail}`,
        "sequence": 1,
        "isEnabled": true,
        "conditions": {
            "fromAddresses": [
                { "emailAddress": { "address": senderEmail } }
            ]
        },
        "actions": {
            "moveToFolder": targetFolderId
        }
    };

    const response = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules", {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(rulePayload)
    });

    if (!response.ok) {
        throw new Error(`Failed to create rule: ${response.statusText}`);
    }
}

// Office.js に関数を登録
Office.actions.associate("moveAndCreateRule", moveAndCreateRule);