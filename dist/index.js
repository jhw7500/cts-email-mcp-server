#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import tls from "tls";
import { simpleParser } from "mailparser";
import nodemailer from "nodemailer";
import fs from "fs";
import path from "path";
import { getTextExtractor } from "office-text-extractor";
class Pop3Client {
    user;
    pass;
    socket = null;
    host = "pop3.hiworks.co.kr";
    port = 995;
    constructor(user, pass) {
        this.user = user;
        this.pass = pass;
    }
    connect() {
        return new Promise((resolve, reject) => {
            this.socket = tls.connect(this.port, this.host, { rejectUnauthorized: false }, () => {
                this.socket?.once('data', (data) => {
                    if (data.toString().startsWith('+OK'))
                        resolve();
                    else
                        reject(new Error('POP3 Greeting failed'));
                });
            });
            this.socket.on('error', (err) => reject(err));
        });
    }
    sendCommand(cmd) {
        return new Promise((resolve, reject) => {
            if (!this.socket)
                return reject(new Error("Not connected"));
            const buffer = [];
            const onData = (data) => {
                buffer.push(data);
                const str = Buffer.concat(buffer).toString();
                if (str.includes('\r\n')) {
                    this.socket?.removeListener('data', onData);
                    resolve(str);
                }
            };
            this.socket.on('data', onData);
            this.socket.write(cmd + "\r\n");
        });
    }
    sendMultiLineCommand(cmd) {
        return new Promise((resolve, reject) => {
            if (!this.socket)
                return reject(new Error("Not connected"));
            const buffer = [];
            const onData = (data) => {
                buffer.push(data);
                const str = Buffer.concat(buffer).toString('binary');
                if (str.endsWith('\r\n.\r\n')) {
                    this.socket?.removeListener('data', onData);
                    resolve(str.slice(0, -5));
                }
            };
            this.socket.on('data', onData);
            this.socket.write(cmd + "\r\n");
        });
    }
    async login() {
        await this.sendCommand(`USER ${this.user}`);
        const res = await this.sendCommand(`PASS ${this.pass}`);
        if (!res.startsWith('+OK'))
            throw new Error(`Login failed: ${res}`);
    }
    async list(count = 10) {
        const statRes = await this.sendCommand("STAT");
        const total = parseInt(statRes.split(' ')[1]);
        const start = Math.max(1, total - count + 1);
        const results = [];
        for (let i = total; i >= start; i--) {
            try {
                const raw = await this.sendMultiLineCommand(`TOP ${i} 0`);
                const parsed = await simpleParser(raw);
                results.push({
                    id: i,
                    date: parsed.date,
                    from: parsed.from?.text,
                    subject: parsed.subject
                });
            }
            catch (e) {
                console.error(`Error fetching ${i}:`, e);
            }
        }
        return results;
    }
    async getEmail(id) {
        const raw = await this.sendMultiLineCommand(`RETR ${id}`);
        const parsed = await simpleParser(raw);
        return {
            id,
            date: parsed.date,
            from: parsed.from?.text,
            subject: parsed.subject,
            body: parsed.text || parsed.html || "(No content)",
            attachments: parsed.attachments.map(a => ({
                filename: a.filename,
                content: a.content,
                contentType: a.contentType
            }))
        };
    }
    quit() {
        if (this.socket) {
            this.socket.write("QUIT\r\n");
            this.socket.end();
        }
    }
}
const server = new McpServer({
    name: "cts-email",
    version: "1.0.0"
});
function getClient() {
    const user = process.env.EMAIL_USER;
    const pass = process.env.EMAIL_PASSWORD;
    if (!user || !pass)
        throw new Error("EMAIL_USER, EMAIL_PASSWORD required in environment");
    return new Pop3Client(user, pass);
}
server.tool("list_emails", "List recent emails in a table format", { count: z.number().default(10) }, async ({ count }) => {
    const client = getClient();
    try {
        await client.connect();
        await client.login();
        const emails = await client.list(count);
        if (emails.length === 0)
            return { content: [{ type: "text", text: "📭 No emails found." }] };
        let table = `| ID | Date | From | Subject |\n|---:|:---|:---|:---|
`;
        emails.forEach(e => {
            const d = e.date ? new Date(e.date).toLocaleString().slice(0, 16) : "Unknown";
            const f = (e.from || "").replace(/\|/g, "&#124;");
            const s = (e.subject || "").replace(/\|/g, "&#124;");
            table += `| ${e.id} | ${d}.. | ${f} | ${s} |
`;
        });
        return { content: [{ type: "text", text: table }] };
    }
    catch (e) {
        return { content: [{ type: "text", text: `Error: ${e.message}` }] };
    }
    finally {
        client.quit();
    }
});
server.tool("search_emails", "Search emails by keyword", { keyword: z.string(), limit: z.number().default(50) }, async ({ keyword, limit }) => {
    const client = getClient();
    try {
        await client.connect();
        await client.login();
        const emails = await client.list(limit);
        const filtered = emails.filter(e => (e.subject?.toLowerCase().includes(keyword.toLowerCase())) ||
            (e.from?.toLowerCase().includes(keyword.toLowerCase())));
        if (filtered.length === 0)
            return { content: [{ type: "text", text: "🔍 No matches found." }] };
        let table = `🔍 **Search Results for '${keyword}'**\n\n| ID | Date | From | Subject |\n|---:|:---|:---|:---|
`;
        filtered.forEach(e => {
            const d = e.date ? new Date(e.date).toLocaleString().slice(0, 16) : "Unknown";
            table += `| ${e.id} | ${d}.. | ${e.from} | ${e.subject} |
`;
        });
        return { content: [{ type: "text", text: table }] };
    }
    catch (e) {
        return { content: [{ type: "text", text: `Error: ${e.message}` }] };
    }
    finally {
        client.quit();
    }
});
server.tool("read_email", "Read specific email content", { id: z.number() }, async ({ id }) => {
    const client = getClient();
    try {
        await client.connect();
        await client.login();
        const email = await client.getEmail(id);
        const attList = email.attachments.map((a) => `\`${a.filename}\``).join(", ") || "(None)";
        const report = `
# 📧 Email Detail (ID: ${email.id})

| Field | Content |
|---|---|
| **From** | ${email.from} |
| **Date** | ${email.date} |
| **Subject** | ${email.subject} |
| **Attachments** | ${attList} |

## 📝 Body
---
${email.body.trim()}
---
`;
        return { content: [{ type: "text", text: report }] };
    }
    catch (e) {
        return { content: [{ type: "text", text: `Error: ${e.message}` }] };
    }
    finally {
        client.quit();
    }
});
server.tool("download_attachment", "Download an attachment", { email_id: z.number(), filename: z.string(), save_path: z.string().default("./downloads") }, async ({ email_id, filename, save_path }) => {
    const client = getClient();
    try {
        await client.connect();
        await client.login();
        const email = await client.getEmail(email_id);
        const attachment = email.attachments.find((a) => a.filename === filename);
        if (!attachment)
            throw new Error("Attachment not found");
        if (!fs.existsSync(save_path))
            fs.mkdirSync(save_path, { recursive: true });
        const filePath = path.join(save_path, filename);
        fs.writeFileSync(filePath, attachment.content);
        return { content: [{ type: "text", text: `✅ Downloaded: acktick${filePath}acktick` }] };
    }
    catch (e) {
        return { content: [{ type: "text", text: `Error: ${e.message}` }] };
    }
    finally {
        client.quit();
    }
});
server.tool("send_email", "Send an email via SMTP", { to: z.string(), subject: z.string(), body: z.string() }, async ({ to, subject, body }) => {
    try {
        const transporter = nodemailer.createTransport({
            host: "smtp.hiworks.co.kr",
            port: 465,
            secure: true,
            auth: {
                user: process.env.EMAIL_USER,
                pass: process.env.EMAIL_PASSWORD
            }
        });
        await transporter.sendMail({
            from: process.env.EMAIL_USER,
            to,
            subject,
            text: body
        });
        return { content: [{ type: "text", text: `✅ Email sent to ${to}` }] };
    }
    catch (e) {
        return { content: [{ type: "text", text: `Error sending email: ${e.message}` }] };
    }
});
server.tool("read_document", "Extract text from local document (pptx, docx, xlsx, pdf)", { file_path: z.string() }, async ({ file_path }) => {
    try {
        if (!fs.existsSync(file_path))
            throw new Error("File not found");
        const extractor = getTextExtractor();
        const text = await extractor.extractText({ input: file_path, type: "file" });
        return { content: [{ type: "text", text: `## 📄 ${path.basename(file_path)}

\`\`\`text
${text}
\`\`\`` }] };
    }
    catch (e) {
        return { content: [{ type: "text", text: `Error reading doc: ${e.message}` }] };
    }
});
async function main() {
    const transport = new StdioServerTransport();
    await server.connect(transport);
}
main().catch((error) => {
    console.error("Server error:", error);
    process.exit(1);
});
