import express from "express";
import cors from "cors";
import passport from "./auth";
import session from "express-session";
import path from "path";
import { google } from 'googleapis';
import { Client } from '@microsoft/microsoft-graph-client';
import cron from 'node-cron';
import { fetchAndSendEmailGoogle, fetchAndSendEmailOutlook } from "./task";

import dotenv from "dotenv";
dotenv.config();

const app = express();

app.use(express.static(path.join(__dirname, 'frontend')));
app.use(express.json());
app.use(cors());

interface RequestWithUser extends express.Request {
    user?: any;
}

function isLoggedin(req: RequestWithUser, res: express.Response, next: express.NextFunction) {
    req.user ? next() : res.sendStatus(401);
}

app.use(session({
    secret: 'keyboard cat',
    resave: false,
    saveUninitialized: true,
    cookie: { secure: false }
}));

app.use(passport.initialize());
app.use(passport.session());

// Routes

app.get('/auth/google',
    passport.authenticate('google', { scope: ['email', 'profile', 'https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.compose', 'https://www.googleapis.com/auth/gmail.modify', 'https://www.googleapis.com/auth/gmail.labels'] })
);

app.get('/auth/google/callback',
    passport.authenticate('google', {
        successRedirect: '/auth/success/google',
        failureRedirect: '/auth/google/failure'
    })
);

app.get("/auth/google/failure", (req, res) => {
    res.send("Something went wrong");
});

app.get(
    "/auth/outlook",
    passport.authenticate("microsoft", {
        scope: ["user.read", "mail.read", "mail.send", "mail.ReadWrite","mailboxSettings.readWrite",""],
    })
);

app.get(
    "/auth/outlook/callback",
    passport.authenticate("microsoft", {
        successRedirect: "/auth/success/outlook",
        failureRedirect: "/auth/outlook/failure",
    })
);

app.get("/auth/outlook/failure", (req, res) => {
    res.send("Outlook authentication failed");
});

app.get("/auth/success/google", isLoggedin, async (req: RequestWithUser, res: express.Response) => {
    try {
        const accessToken = req.user.tokens.access_token;

        const oauth2Client = new google.auth.OAuth2();
        oauth2Client.setCredentials({ access_token: accessToken });
        //Hit First Time
        const response = await fetchAndSendEmailGoogle(oauth2Client)

        //Then Go To Background Task
        let cronjob:any;
        cronjob = cron.schedule('*/30 * * * * *', async () => {
            try {
                const GoogleResponse = await fetchAndSendEmailGoogle(oauth2Client);
                if (GoogleResponse.message == "No New Mails") {
                    cronjob.stop();
                    console.log('Stopping Background Task No Emails left to analyze:', GoogleResponse);
                } else {
                    console.log('Background task executed:', GoogleResponse);
                }
            } catch (error) {
                console.error('Error executing background task:', error);
            }
        });

        res.json(response);

    } catch (error) {
        console.error('Error fetching and analyzing emails:', error);
        res.status(500).send('Error fetching and analyzing emails');
    }
});

app.get("/auth/success/outlook", isLoggedin, async (req: RequestWithUser, res: express.Response) => {
    try {
        const accessToken = req.user.accessToken;
        const graph = Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            },
        });
        //Hit First time:
        const response = await fetchAndSendEmailOutlook(graph)

        //Then Go To Background Task
        let cronjob:any;
        cronjob = cron.schedule('*/30 * * * * *', async () => {
            try {
                const outlookResponse = await fetchAndSendEmailOutlook(graph);
                if (outlookResponse.message == "No New Mails") {
                    cronjob.stop()
                    console.log('Background task Is stopped no new mails to analyze', outlookResponse);
                }
                console.log('Background task executed:', outlookResponse);
            } catch (error) {
                console.error('Error executing background task:', error);
            }
        });
        res.json(response);

    } catch (error) {
        console.error('Error fetching and analyzing emails:', error);
        res.status(500).send('Error fetching and analyzing emails');
    }
});

app.get("*", (req, res) => {
    res.sendFile('index.html', { root: path.join(__dirname) });
});

app.listen(5500, () => {
    console.log("Server started");
});