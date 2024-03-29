import { google, gmail_v1 } from 'googleapis';
import { Client } from '@microsoft/microsoft-graph-client';
import { OpenAI } from 'openai';

interface EmailReply {
    subject: string;
    body: string;
}

interface AnalyzedResponse {
    label: string;
    extractedMailContent: string;
    replyMail: EmailReply;
}

const analyzeEmail = async (emailContent: string): Promise<string> => {
    try {
        const openai = new OpenAI({
            apiKey: process.env.OPENAI_API_KEY || '',
            dangerouslyAllowBrowser: true,
        });
        const response = await openai.chat.completions.create({
            messages: [
                {
                    role: 'system',
                    content:`Consider yourself a recruiter and you need to send replies to applicants finding job at your comapny You need to categorize the mails in 4 labels ["Interested","More Information","Not Interested","Others"] .\n
                    1.Others are those kind of mail which is not related to jobs or recruitment process in this case you need to ignore the mail no need to generate the reply just assing the label as Others\n
                    2.Interested are those kind of mail which says that applicant is interested for the job role and whatever skills mentionend in the mail is alligning to the job profile in this kind of mail assign the label as Interested and generate reply asking them to weather they want to come on a call and also assign the label as Interested.\n
                    3.Not Interested are those kind of mail where the applicant wants to withdraw from a recruitment process going on or he/she is not interested in a job role applied in these case generate appropriate response wishing them good luck etc also assing the label as Not Interested.\n
                    4.More Information are those kind of mail where the candidate is interested for a job role but has not mentioned any more details about them or skills for example this mail : egarding any role I want to join your company I want to join your company. It says that he candidate is only intersted for a role in theior company but hasn't mentionend about their skills or any other profile details \n
                    Remeber to give your resposne in following json format: \n
                    {
                        "label":"",
                        "extractedMailContent":"",
                        "replyMail":{
                            "subject":"",
                            "body":""
                        }
                    }
                            This is the email content below: \n
                    ${emailContent}`
                },
            ],
            model: 'gpt-3.5-turbo',
        });
        console.log(response.choices[0].message.content)
        return response.choices[0].message.content || "";
    } catch (error) {
        throw error;
    }
};

const sendReplyUsingGoogle = async (gmail: gmail_v1.Gmail, message: gmail_v1.Schema$Message, replyMail: EmailReply): Promise<void> => {
    try {
        const messageId = message.id ?? '';
        const res = await gmail.users.messages.get({
            userId: 'me',
            id: messageId,
            format: 'metadata',
            metadataHeaders: ['Subject', 'From'],
        });
        const from = res.data.payload?.headers?.find((header) => header.name == 'From')?.value || '';
        const replyTo = from.match(/<(.*)>/)?.[1] || '';
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

        await gmail.users.messages.send({
            userId: 'me',
            requestBody: {
                raw: encodedMessage,
            },
        });
    } catch (error) {
        console.error('Error sending reply email:', error);
        throw error;
    }
};


const sendReplyUsingOutlook = async (client: Client, message: any, replyMail: EmailReply): Promise<void> => {
    try {
        const fromEmailAddress = message.sender?.emailAddress?.address || '';

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

        await client.api(`/me/sendMail`).post(replyEmail);
    } catch (error) {
        console.error('Error sending reply email:', error);
        throw error;
    }
};

async function getGmailLabelId(gmail: gmail_v1.Gmail, labelName: string) {
    let labelId = null;
    const labels = await gmail.users.labels.list({ userId: 'me' });
    const label = labels.data.labels.find(label => label.name === labelName);

    if (label) {
        labelId = label.id;
    } else {
        const createdLabel = await gmail.users.labels.create({
            userId: 'me',
            requestBody: {
                name: labelName,
                labelListVisibility: 'labelShow',
                messageListVisibility: 'show'
            }
        });
        labelId = createdLabel.data.id;
    }

    return labelId;
}

async function addOutlookCategory(graph: Client, message: any, category: string) {
    await graph.api(`/me/messages/${message.id}`).post({
        categories: [category]
    });
}

const fetchAndSendEmailGoogle = async (oauth2Client: any): Promise<any> => {
    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

    const messages = await (gmail.users.messages.list as any)({ userId: 'me', maxResults: 1, q: 'is:unread', orderBy: 'date desc' });

    const messageContents = await Promise.all(messages.data.messages.map(async (message: any) => {
        const messageData = await gmail.users.messages.get({ userId: 'me', id: message.id });
        const payload = messageData.data.payload;
        const subject = payload.headers.find((header)=> header.name === "Subject")?.value;
        let content = ""
        if (payload.parts){
            const parts = payload.parts.find(
                (part)=>part.mimeType = "text/plain"
            )
            if(parts){
                content = Buffer.from(parts.body.data,"base64").toString("utf-8");
            }
        }else{
            content = Buffer.from(payload.body.data, "base64").toString("utf-8");
        }
        const snippet = messageData.data.snippet
        const body = `${subject} ${snippet} ${content} `
        return { id: message.id, body: body };
    }));

    const sentEmail: any[] = [];
    const messageReplied: string[] = [];
    for (const message of messageContents) {
        const analyzedResponse = await analyzeEmail(message.body || "");
        console.log(analyzedResponse)
        const response: AnalyzedResponse = JSON.parse(analyzedResponse);
        const replyMail = response.replyMail;
        const extractedMailContent = response.extractedMailContent;
        messageReplied.push(extractedMailContent);
        if (response.label == "Interested") {
            const labelId = await getGmailLabelId(gmail, 'Interested');
            await (gmail.users.messages.modify as any)({
                userId: 'me',
                id: message.id,
                addLabelIds: [labelId],
            });
        }
        else if (response.label == "Not Interested") {
            const labelId = await getGmailLabelId(gmail, 'Not Interested');
            await (gmail.users.messages.modify as any)({
                userId: 'me',
                id: message.id,
                addLabelIds: [labelId],
            });
        }
        else if (response.label == "More Information") {
            const labelId = await getGmailLabelId(gmail, 'More Information');
            await (gmail.users.messages.modify as any)({
                userId: 'me',
                id: message.id,
                addLabelIds: [labelId],
            });
        }
        else if (response.label == "Others") {
            const labelId = await getGmailLabelId(gmail, 'Others');
            await (gmail.users.messages.modify as any)({
                userId: 'me',
                id: message.id,
                addLabelIds: [labelId],
            });
        }
        if (replyMail?.body) {
            await sendReplyUsingGoogle(gmail, message, replyMail);
            sentEmail.push([`${response.label.toUpperCase()} - Mail is of ${response.label.toUpperCase()} category`, replyMail]);
        } else {
            sentEmail.push(["OTHER", "Mail is in other category no reply sent"]);
        }

        // Mark the email as read
        await gmail.users.messages.modify({
            userId: 'me',
            id: message.id,
            requestBody: {
                removeLabelIds: ['UNREAD']
            }
        });
    }
    return { message: messageReplied.length>0?"Success":"No New Mails", email: messageReplied.length > 0 ? messageReplied : ["No Mails"], replies: sentEmail.length > 0 ? sentEmail : ["No Mails"] };
};

const fetchAndSendEmailOutlook = async (graph: Client): Promise<any> => {
    const messages = await graph.api("/me/messages").filter("isRead eq false").orderby("receivedDateTime desc").top(1).get();
    const messageContents = messages.value ? await Promise.all(messages.value.map(async (message: any) => {
        const messageData = await graph.api(`/me/messages/${message.id}`).get();
        return messageData;
    })) : [];

    const sentEmail: any[] = [];
    const messageReplied: string[] = [];
    for (const message of messageContents) {
        const analyzedResponse = await analyzeEmail(message.body.content);
        const response: AnalyzedResponse = JSON.parse(analyzedResponse);
        const replyMail = response.replyMail;
        const extractedMailContent = response.extractedMailContent;
        messageReplied.push(extractedMailContent);
        if (response.label === "Interested" || response.label === "Not Interested" || response.label === "More Information" || response.label === "Others") {
            await addOutlookCategory(graph, message, response.label);
        }
        if (replyMail.body) {
            await sendReplyUsingOutlook(graph, message, replyMail);
            sentEmail.push([`${response.label.toUpperCase()} - Mail is of ${response.label.toUpperCase()} category`, replyMail]);
        } else {
            sentEmail.push([response.label.toUpperCase(), "Mail is in other category no reply sent"]);
        }

        await graph.api(`/me/messages/${message.id}`).update({
            isRead: true
        });
    }
    return { message: "Success", email: messageReplied.length > 0 ? messageReplied : ["No Mails"], replies: sentEmail.length > 0 ? sentEmail : ["No Mails"] };
};


export { fetchAndSendEmailGoogle, fetchAndSendEmailOutlook, sendReplyUsingGoogle, sendReplyUsingOutlook };
