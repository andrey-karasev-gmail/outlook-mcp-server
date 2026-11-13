#!/usr/bin/env node

/**
 * Script to get Outlook authentication token
 * 
 * Usage:
 *   node scripts/get-auth-token.js
 *
 * This script starts a temporary HTTP server at http://localhost:3000/callback,
 * opens the authorization URL, and captures the authorization code automatically.
 */

import http from "http";
import https from "https";
import { exec } from "child_process";
import dotenv from "dotenv";

dotenv.config();

const CLIENT_ID = process.env.OUTLOOK_CLIENT_ID;
const CLIENT_SECRET = process.env.OUTLOOK_CLIENT_SECRET;
const TENANT_ID = process.env.OUTLOOK_TENANT_ID || "common";
const REDIRECT_URI = "http://localhost:3000/callback";
const PORT = 3000;

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error("Error: OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET must be set in .env");
  process.exit(1);
}

function openBrowser(url) {
  const platform = process.platform;
  if (platform === "win32") {
    exec(`start "" "${url}"`, (err) => {
      if (err) {
        console.warn("Could not open browser automatically. Open the URL manually.");
      }
    });
  } else if (platform === "darwin") {
    exec(`open "${url}"`, (err) => {
      if (err) {
        console.warn("Could not open browser automatically. Open the URL manually.");
      }
    });
  } else {
    exec(`xdg-open "${url}"`, (err) => {
      if (err) {
        console.warn("Could not open browser automatically. Open the URL manually.");
      }
    });
  }
}

function getAuthUrl() {
  return `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
    `client_id=${CLIENT_ID}&` +
    `redirect_uri=${encodeURIComponent(REDIRECT_URI)}&` +
    `response_type=code&` +
    `response_mode=query&` +
    `scope=${encodeURIComponent("openid profile offline_access Mail.Read Mail.ReadWrite MailboxSettings.ReadWrite")}&` +
    `prompt=select_account`;
}

async function getAuthCode() {
  return new Promise((resolve, reject) => {
    const authUrl = getAuthUrl();

    const server = http.createServer((req, res) => {
      if (!req.url) {
        res.writeHead(400);
        res.end("Invalid request");
        return;
      }

      const requestUrl = new URL(req.url, REDIRECT_URI);
      if (requestUrl.pathname !== "/callback") {
        res.writeHead(404);
        res.end("Not found");
        return;
      }

      const code = requestUrl.searchParams.get("code");
      const error = requestUrl.searchParams.get("error");

      if (error) {
        res.writeHead(400, { "Content-Type": "text/html" });
        res.end(`<h1>Authentication failed</h1><p>${error}</p>`);
        server.close();
        reject(new Error(`Authorization error: ${error}`));
        return;
      }

      if (!code) {
        res.writeHead(400, { "Content-Type": "text/html" });
        res.end("Missing authorization code");
        server.close();
        reject(new Error("Authorization code missing"));
        return;
      }

      res.writeHead(200, { "Content-Type": "text/html" });
      res.end("<h1>Authentication successful</h1><p>You may close this window and return to the terminal.</p>");
      server.close();
      resolve(code);
    });

    server.listen(PORT, () => {
      console.log("\n1. Your browser should open automatically.");
      console.log("2. If it does not, open this URL manually:");
      console.log(authUrl);
      openBrowser(authUrl);
    });

    server.on("error", (err) => {
      reject(err);
    });
  });
}

async function exchangeCodeForToken(code) {
  return new Promise((resolve, reject) => {
    const data = new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      code: code,
      redirect_uri: REDIRECT_URI,
      grant_type: "authorization_code"
    }).toString();

    const options = {
      hostname: "login.microsoftonline.com",
      path: `/${TENANT_ID}/oauth2/v2.0/token`,
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Content-Length": Buffer.byteLength(data)
      }
    };

    const req = https.request(options, (res) => {
      let body = "";
      res.on("data", (chunk) => {
        body += chunk;
      });
      res.on("end", () => {
        if (res.statusCode === 200) {
          resolve(JSON.parse(body));
        } else {
          reject(new Error(`Failed to get token: ${body}`));
        }
      });
    });

    req.on("error", reject);
    req.write(data);
    req.end();
  });
}

async function main() {
  try {
    console.log("Outlook MCP Server - Authentication Token Generator\n");

    const code = await getAuthCode();
    console.log("\nExchanging code for tokens...");

    const tokens = await exchangeCodeForToken(code);

    console.log("\n✓ Success!\n");
    console.log("Add this to your .env file:\n");
    console.log(`OUTLOOK_REFRESH_TOKEN=${tokens.refresh_token}\n`);

    if (tokens.access_token) {
      console.log("(Access token expires in", tokens.expires_in, "seconds)\n");
    }
  } catch (error) {
    console.error("Error:", error.message);
    process.exit(1);
  }
}

main();
