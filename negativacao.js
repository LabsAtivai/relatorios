const fs = require("fs");
const { google } = require("googleapis");
const axios = require("axios");

const credentials = {
  type: "service_account",
  project_id: "ativa2-411320",
  private_key_id: "54d93f53c9572831254eae9f88dd08a4e4e1ad63",
  private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC+QxgN5pmYFAuA\nOxx11HixRGGTWin09pTFyyBMWoS0fWAWDlapjwEndWeDXpMqZ83nvifFzBWLF5bm\nYUKHJWR//0XLLIq5+6CVRZxqv+/lucITSjflRV+TqzepF7NpUoU4mKG2vaCLMblo\nh6KgU/U2e4LWzhuhml8A0RCfydnGENdctZI+Tn4lRtdLzzABGp5Va/W4vVdSPdFo\nAWruPyu9M0ICCejZfWS0dWyyxvWRNw2a/T6b6lgUOV76g71guBVlOKLIOWHYJjJE\nexHHXJG5w7cglJCWRB5mmDQ+dGycqE8xO/9X+i3rZFRMW2AY9lFQTuddquBRVx3i\n6GQwsSSnAgMBAAECggEACuTblTwtcjw/llWMISEL0haYwa+pdFnEFfk1/bk3HQCK\nxmiGxMmn5sL9rNN2+Lgd3ZWRQG2ZlC/DF6jc/tEAyqVqsSH5RYHnZXSpaqFX0p8K\nwkD/J1UMDnAAzWFKiA7OCvoOVGDSNvYfaCFQCf8UrMxwpu0BPwUQSIMwAP41RnZp\n7pSjuBCztU6zH6k8kfGa6ZiQ27gkTX0K2DAWfdrHQbD4MjM7Te9YV6MvVhkx2uMy\nF1a16ufp78o3Um1qL7xvdDv6EZUCnyjzlRI9aVQQLrrR/JbbHPouVjEo0yS7q30X\nlMWlt1kPXoyNVTFiHNiqbStG/9Kn7equIHBtbU85AQKBgQD5089PFdm69t+boZe0\nJJO3opiB+VsNEcW4ViF/5hlT0Kv6Ne7bcb7V97qeLMoT+1ZD0cUhMa5OliGZeJjp\nElvoEkI+rvOpJtkQjRBBzdktUSDwte0ScozaG3Op9VVB+AsRTqcFtbSsjQ66dniJ\n8NzTFJJXiI6Rcz8cM+iU6Zj/pwKBgQDC9oap/MzIgHrIe6GWC8rTPeV3NRl/HUcO\n7FuOhx+GgQfRHW4lPLiARoje4IygXEMHN4MkBkSTUc6QAyVsbdX5FtSIWrDqyhg+\nZuTt9AXuxoIMWXkoH70uO43w1eRoIRmarNeYCzYBGu/s4MlnP1gRQuLJ0EpXFU+0\nyki2kMNTAQKBgHhA5lcRE38Vekj1nYpO2PPZxkM5/gPqfBvhCbsAzUw087M65sCm\nnc9ssV2V/adkof9/J750pYyaY432KNR3y9mHgG+f0xWm53E6FYs3RMS1en+jcnwg\nu3/5GtHCj8lzVhB8pZTwBAnS1DYY9KihUgogqtiLmOSWbthqfBfB4a2dAoGBAJ5K\nvodho3rfJdGE32sN4/2i8Z3Z1Pup77mkGaoc93GjbY9RT86YAEzV9+bNdh/1CM7h\nOW6UUDU0ZHv0sfvZKbN139VdnOrkbs6riA/S4sY9EfWo53+2VUkmPVZes3X/+ePl\nlZ3y7EP1dPtkfuF/QqexIKUuok0WFoL5AhSIcK4BAoGBAIL1j8sd9YYHoqCfnzZq\nLigR5dS/BRUcxNwA4V84OnsOEWLAJ1hH5kw9U2q0rrlQZP9aatYWcZNSz1qRsfvG\nvCV4Muq5R6GoO617Ct/RHD8My638vqJXB4QAR6KbCDdhwVjSn/REao1rmMc0vB+a\nWhKNBp03AOMjavuL/V+COgVX\n-----END PRIVATE KEY-----\n",
  client_email: "negativacao@ativa2-411320.iam.gserviceaccount.com",
  client_id: "110280029940213741854",
  auth_uri: "https://accounts.google.com/o/oauth2/auth",
  token_uri: "https://oauth2.googleapis.com/token",
  auth_provider_x509_cert_url: "https://www.googleapis.com/oauth2/v1/certs",
  client_x509_cert_url:
    "https://www.googleapis.com/robot/v1/metadata/x509/negativacao%40ativa2-411320.iam.gserviceaccount.com",
  universe_domain: "googleapis.com",
};

async function getGmailAccessToken(clientId, clientSecret, refreshToken, userEmail) {
  const payload = {
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: refreshToken,
    grant_type: "refresh_token",
  };
  const params = new URLSearchParams(payload);
  const tokenUrl = "https://accounts.google.com/o/oauth2/token?" + params.toString();

  try {
    const response = await axios.post(tokenUrl, {}, {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });
    return response.data.access_token || null;
  } catch (error) {
    console.error(`❌ Failed to get token for ${userEmail}:`, error.response?.data || error.message);
    return null;
  }
}

async function getSnovAccessToken(clientIdValue, clientSecretValue) {
  const response = await axios.get('https://api.snov.io/v1/oauth/access_token', {
    params: { grant_type: 'client_credentials', client_id: clientIdValue, client_secret: clientSecretValue }
  });
  return response.data.access_token;
}

async function addToDoNotEmailList(emails, listId, clientIdValue, clientSecretValue) {
  const token = await getSnovAccessToken(clientIdValue, clientSecretValue);
  const params = new URLSearchParams();
  params.append('access_token', token);
  params.append('listId', listId);
  emails.forEach(email => params.append('items[]', email));

  try {
    const response = await axios.post('https://api.snov.io/v1/do-not-email-list', params, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });
    console.log('✅ Emails negativados com sucesso:', response.data);
  } catch (error) {
    console.error('❌ Erro ao negativar emails:', error.response?.data || error.message);
  }
}

async function loadUsersFromSpreadsheet(spreadsheetId) {
  const client = await google.auth.getClient({
    credentials,
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/gmail.readonly"
    ]
  });

  const sheets = google.sheets({ version: "v4", auth: client });
  const range = "Todos6.0";
  const response = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  const rows = response.data.values || [];

  const users = [];
  for (const data of rows) {
    const [userEmail, refreshToken, googleClientId, googleClientSecret, snovId, snovSecret, listId, folderId, snovEmail, snovPass] = data;
    const gmailToken = await getGmailAccessToken(googleClientId, googleClientSecret, refreshToken, userEmail);
    if (gmailToken) {
      users.push({
        USER_ID: userEmail,
        GMAIL_TOKEN: gmailToken,
        GOOGLE_CLIENT_ID: googleClientId,
        GOOGLE_CLIENT_SECRET: googleClientSecret,
        SNOVIO_ID: snovId,
        SNOVIO_SECRET: snovSecret,
        LIST_ID: listId,
        FOLDER_ID: folderId,
        SNOVIO_EMAIL: snovEmail,
        SNOVIO_PASS: snovPass
      });
    }
  }
  return users;
}

async function processUser(user, startTimestamp, endTimestamp) {
  if (!user.GMAIL_TOKEN) throw new Error(`Token ausente: ${user.USER_ID}`);
  const url = `https://gmail.googleapis.com/gmail/v1/users/${user.USER_ID}/messages?q=after:${startTimestamp}&before:${endTimestamp}`;

  try {
    const res = await axios.get(url, {
      headers: { Authorization: `Bearer ${user.GMAIL_TOKEN}` },
    });

    const emailList = [];
    const messages = res.data.messages || [];
    console.log(`📬 ${messages.length} mensagens encontradas para ${user.USER_ID}`);

    for (const msg of messages) {
      try {
        const msgData = await axios.get(`https://gmail.googleapis.com/gmail/v1/users/${user.USER_ID}/messages/${msg.id}`, {
          headers: { Authorization: `Bearer ${user.GMAIL_TOKEN}` },
        });
        const headers = msgData.data.payload.headers || [];
        for (const h of headers) {
          if (h.name === "From") {
            const match = h.value.match(/[\w.-]+@[\w.-]+\.[A-Za-z]{2,}/);
            if (match) emailList.push(match[0]);
          }
        }
      } catch (err) {
        console.error(`❗ Erro na mensagem ${msg.id} de ${user.USER_ID}:`, err);
      }
    }

    const uniqueEmails = [...new Set(emailList)];
    if (uniqueEmails.length) {
      console.log(`📧 Emails únicos para ${user.USER_ID}:`, uniqueEmails);
      await addToDoNotEmailList(uniqueEmails, user.LIST_ID, user.SNOVIO_ID, user.SNOVIO_SECRET);
      fs.appendFileSync("log_negativados.txt", `Usuário: ${user.USER_ID}\nEmails:\n${uniqueEmails.join('\n')}\n\n`);
    } else {
      console.log(`⚠️ Nenhum email válido para negativar: ${user.USER_ID}`);
    }
  } catch (err) {
    console.error(`❌ Erro ao processar usuário ${user.USER_ID}:`, err);
  }
}

async function fetchAndNegateEmails() {
  const spreadsheetId = "1d6SL3vBavlh4flFvnI-PweHrd3AJe9PMHh50PshkR8c";
  const users = await loadUsersFromSpreadsheet(spreadsheetId);

  console.log(`🚀 ${users.length} usuários carregados.`);
  const threeDaysAgo = new Date();
  threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
  const start = Math.floor(threeDaysAgo.getTime() / 1000);
  const end = Math.floor(Date.now() / 1000);

  for (const user of users) {
    await processUser(user, start, end);
  }
}

fetchAndNegateEmails();