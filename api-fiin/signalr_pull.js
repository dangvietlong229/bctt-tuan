const fs = require('fs');
const path = require('path');

const HUB_URL = "https://realtime.fiintrade.vn/RealtimeHub";
const REST_URL = "https://technical.fiintrade.vn/TradingView/GetStockEvents?OrganCode=ACB&From=2025-03-28T02%3A00%3A00.000Z&To=2036-12-31T17%3A00%3A00.000Z&language=vi";

const browserHeaders = {
  "accept": "*/*",
  "accept-language": "vi,en-US;q=0.9,en;q=0.8,zh-CN;q=0.7,zh;q=0.6",
  "origin": "https://fiintrade.vn",
  "referer": "https://fiintrade.vn/",
  "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36"
};

const RECORD_SEPARATOR = String.fromCharCode(0x1e);

async function run() {
  try {
    console.log("1. Negotiating connection...");
    const negotiateRes = await fetch(`${HUB_URL}/negotiate`, {
      method: "POST",
      headers: {
        ...browserHeaders,
        "content-type": "text/plain;charset=UTF-8",
        "x-requested-with": "XMLHttpRequest"
      },
      body: ""
    });

    if (!negotiateRes.ok) {
      throw new Error(`Negotiate failed: ${negotiateRes.status} ${negotiateRes.statusText}`);
    }

    const negotiateData = await negotiateRes.json();
    const connectionId = negotiateData.connectionId;
    console.log(`Negotiated Connection ID: ${connectionId}`);

    // Build the WebSocket URL
    const wsUrl = `wss://realtime.fiintrade.vn/RealtimeHub?id=${connectionId}`;
    console.log(`2. Connecting to WebSocket: ${wsUrl}`);

    const ws = new WebSocket(wsUrl, {
      headers: browserHeaders
    });

    ws.onopen = () => {
      console.log("WebSocket connected. Sending SignalR handshake...");
      // Send SignalR protocol handshake
      const handshake = JSON.stringify({ protocol: "json", version: 1 }) + RECORD_SEPARATOR;
      ws.send(handshake);
    };

    ws.onmessage = async (event) => {
      const messages = event.data.split(RECORD_SEPARATOR);
      for (const rawMessage of messages) {
        if (!rawMessage.trim()) continue;
        try {
          const msg = JSON.parse(rawMessage);
          console.log("Received Message from Hub:", JSON.stringify(msg, null, 2));

          // SignalR ping (type 6) - reply to keep-alive if necessary
          if (msg.type === 6) {
            ws.send(JSON.stringify({ type: 6 }) + RECORD_SEPARATOR);
            continue;
          }

          // Look for any string payload that looks like a u0 token
          // Typical token is base64 string ending with == or similar length
          // Let's inspect target message properties
          if (msg.arguments && msg.arguments.length > 0) {
            const potentialToken = msg.arguments[0];
            if (typeof potentialToken === 'string' && potentialToken.length > 20 && potentialToken.endsWith('==')) {
              console.log(`\n[SUCCESS] Captured Live Token: ${potentialToken}`);
              ws.close();
              
              // 3. Query the target REST API using the newly captured token
              console.log("\n3. Fetching REST API with captured token...");
              const apiHeaders = {
                "accept": "application/json, text/plain, */*",
                "accept-language": "vi,en-US;q=0.9,en;q=0.8,zh-CN;q=0.7,zh;q=0.6",
                "authorization": "Bearer",
                "sec-ch-ua": "\"Google Chrome\";v=\"149\", \"Chromium\";v=\"149\", \"Not)A;Brand\";v=\"24\"",
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": "\"Windows\"",
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-site",
                "u0": potentialToken,
                ...browserHeaders
              };

              const apiRes = await fetch(REST_URL, { headers: apiHeaders });
              console.log(`REST API Status: ${apiRes.status} ${apiRes.statusText}`);
              
              if (apiRes.ok) {
                const apiData = await apiRes.json();
                const outPath = path.join(__dirname, 'responses', 'GetStockEvents_Live.json');
                fs.writeFileSync(outPath, JSON.stringify(apiData, null, 2));
                console.log(`Saved live data successfully to: ${outPath}`);
              } else {
                const text = await apiRes.text();
                console.log("Failed body:", text);
              }
              return;
            }
          }
        } catch (e) {
          // Message might not be valid JSON or not the handshake response
        }
      }
    };

    ws.onerror = (err) => {
      console.error("WebSocket Error:", err);
    };

    ws.onclose = () => {
      console.log("WebSocket connection closed.");
    };

    // Close connection after 15 seconds if no token is captured to prevent hanging
    setTimeout(() => {
      console.log("Timeout reached. Closing socket.");
      ws.close();
    }, 15000);

  } catch (error) {
    console.error("Error in execution loop:", error.message);
  }
}

run();
