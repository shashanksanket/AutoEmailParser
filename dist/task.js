"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.sendReplyUsingOutlook = exports.sendReplyUsingGoogle = exports.fetchAndSendEmailOutlook = exports.fetchAndSendEmailGoogle = void 0;
const googleapis_1 = require("googleapis"); // Import Google APIs
const openai_1 = require("openai"); // Import OpenAI object
const analyzeEmail = (emailContent) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        const openai = new openai_1.OpenAI({
            apiKey: process.env.OPENAI_API_KEY || '',
            dangerouslyAllowBrowser: true,
        });
        const response = yield openai.chat.completions.create({
            messages: [
                {
                    role: 'system',
                    content: `analyze this email content and make a reply mail return your response in following json format {label:"",extractedMailContent:"",replyMail:{subject:"",body:""}} remeber label can be anything drom following [interested, not interested, more information,other] I basically want to read those mails which are related to jobs I want to reply people contacting for jobs with me so if there is any other mail ignore them and assign them label as other I repeat do not generate body of the mails if it is in other category that is not related to jobs${emailContent}`,
                },
            ],
            model: 'gpt-3.5-turbo',
        });
        return response.choices[0].message.content || "";
    }
    catch (error) {
        throw error;
    }
});
const sendReplyUsingGoogle = (gmail, message, replyMail) => __awaiter(void 0, void 0, void 0, function* () {
    var _a, _b, _c, _d, _e, _f, _g, _h;
    try {
        const messageId = (_a = message.id) !== null && _a !== void 0 ? _a : ''; // Ensure messageId is a string
        const res = yield gmail.users.messages.get({
            userId: 'me',
            id: messageId,
            format: 'metadata',
            metadataHeaders: ['Subject', 'From'],
        });
        const subject = ((_d = (_c = (_b = res.data.payload) === null || _b === void 0 ? void 0 : _b.headers) === null || _c === void 0 ? void 0 : _c.find((header) => header.name == 'Subject')) === null || _d === void 0 ? void 0 : _d.value) || '';
        const from = ((_g = (_f = (_e = res.data.payload) === null || _e === void 0 ? void 0 : _e.headers) === null || _f === void 0 ? void 0 : _f.find((header) => header.name == 'From')) === null || _g === void 0 ? void 0 : _g.value) || '';
        const replyTo = ((_h = from.match(/<(.*)>/)) === null || _h === void 0 ? void 0 : _h[1]) || '';
        const replySubject = replyMail.subject;
        const replyBody = replyMail.body.split('.').join('\n');
        const rawMessage = [
            `From: me`,
            `To: ${replyTo}`,
            `Subject: ${replySubject}`,
            `In-Reply-To: ${messageId}`,
            `References: ${messageId}`,
            ``,
            replyBody,
        ].join('\n');
        const encodedMessage = Buffer.from(rawMessage).toString('base64').replace(/\+/g, '-').replace(/\//g, '-').replace(/=+$/, '');
        yield gmail.users.messages.send({
            userId: 'me',
            requestBody: {
                raw: encodedMessage,
            },
        });
    }
    catch (error) {
        console.error('Error sending reply email:', error);
        throw error;
    }
});
exports.sendReplyUsingGoogle = sendReplyUsingGoogle;
const sendReplyUsingOutlook = (client, message, replyMail) => __awaiter(void 0, void 0, void 0, function* () {
    var _j, _k;
    try {
        const messageId = message.id;
        const fromEmailAddress = ((_k = (_j = message.sender) === null || _j === void 0 ? void 0 : _j.emailAddress) === null || _k === void 0 ? void 0 : _k.address) || '';
        const replyEmail = {
            message: {
                subject: replyMail.subject,
                body: {
                    contentType: 'Text',
                    content: replyMail.body,
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: fromEmailAddress,
                        },
                    },
                ],
            },
            saveToSentItems: 'false',
        };
        yield client.api(`/me/sendMail`).post(replyEmail);
    }
    catch (error) {
        console.error('Error sending reply email:', error);
        throw error;
    }
});
exports.sendReplyUsingOutlook = sendReplyUsingOutlook;
const fetchAndSendEmailGoogle = (oauth2Client) => __awaiter(void 0, void 0, void 0, function* () {
    var _l;
    const gmail = googleapis_1.google.gmail({ version: 'v1', auth: oauth2Client });
    const messages = yield gmail.users.messages.list({ userId: 'me', maxResults: 1 });
    const messageContents = ((_l = messages.data) === null || _l === void 0 ? void 0 : _l.messages) ? yield Promise.all(messages.data.messages.map((message) => __awaiter(void 0, void 0, void 0, function* () {
        const messageData = yield gmail.users.messages.get({ userId: 'me', id: message.id });
        return messageData.data;
    }))) : [];
    const sentEmail = [];
    const messageReplied = [];
    for (const message of messageContents) {
        const analyzedResponse = yield analyzeEmail(message.snippet || "");
        const response = JSON.parse(analyzedResponse);
        const replyMail = response.replyMail;
        const extractedMailContent = response.extractedMailContent;
        messageReplied.push(extractedMailContent);
        if (replyMail.body) {
            yield sendReplyUsingGoogle(gmail, message, replyMail);
            sentEmail.push([`${response.label.toUpperCase()} - Mail is of ${response.label.toUpperCase()} category`, replyMail]);
        }
        else {
            sentEmail.push([response.label.toUpperCase(), "Mail is in other category no reply sent"]);
        }
    }
    return { message: "Success", email: messageReplied, replies: sentEmail };
});
exports.fetchAndSendEmailGoogle = fetchAndSendEmailGoogle;
const fetchAndSendEmailOutlook = (graph) => __awaiter(void 0, void 0, void 0, function* () {
    const messages = yield graph.api("/me/messages").top(1).get();
    const messageContents = messages.value ? yield Promise.all(messages.value.map((message) => __awaiter(void 0, void 0, void 0, function* () {
        const messageData = yield graph.api(`/me/messages/${message.id}`).get();
        return messageData;
    }))) : [];
    const sentEmail = [];
    const messageReplied = [];
    for (const message of messageContents) {
        const analyzedResponse = yield analyzeEmail(message.body.content);
        const response = JSON.parse(analyzedResponse);
        const replyMail = response.replyMail;
        const extractedMailContent = response.extractedMailContent;
        messageReplied.push(extractedMailContent);
        if (replyMail.body) {
            yield sendReplyUsingOutlook(graph, message, replyMail);
            sentEmail.push([`${response.label.toUpperCase()} - Mail is of ${response.label.toUpperCase()} category`, replyMail]);
        }
        else {
            sentEmail.push([response.label.toUpperCase(), "Mail is in other category no reply sent"]);
        }
    }
    console.log(sentEmail);
    return { message: "Success", email: messageReplied, replies: sentEmail };
});
exports.fetchAndSendEmailOutlook = fetchAndSendEmailOutlook;
//# sourceMappingURL=task.js.map