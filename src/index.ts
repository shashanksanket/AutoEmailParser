import express from "express";
import cors from "cors";
import passport from "./auth";
import session from "express-session";
import path from "path";
import { google } from 'googleapis';
import { Client } from '@microsoft/microsoft-graph-client';
import cron from 'node-cron';
import { emailQueue } from './queue';

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
        scope: ["user.read", "mail.read", "mail.send", "mail.ReadWrite", "mailboxSettings.readWrite", ""],
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
        const provider = "google"
        emailQueue.add('sendEmail', { accessToken, provider },{repeat: {every: 10000}});
        res.json("Job Started see terminal for updates");
    } catch (error) {
        console.error('Error fetching and analyzing emails:', error);
        res.status(500).send('Error fetching and analyzing emails');
    }
});

app.get("/auth/success/outlook", isLoggedin, async (req: RequestWithUser, res: express.Response) => {
    try {
        const accessToken = req.user.accessToken;
        const provider = "outlook"
        emailQueue.add('sendEmail', { accessToken, provider },{repeat: {every: 10000}, removeOnComplete: true, removeOnFail: true});
        res.json("Job Started see terminal for updates");
    } catch (error) {
        console.error('Error fetching and analyzing emails:', error);
        res.status(500).send('Error fetching and analyzing emails');
    }
});

app.get("*", (req, res) => {
    res.sendFile('index.html', { root: path.join(__dirname) });
});

app.listen(5500, async () => {
    await emailQueue.clean(0, 999999);
    await emailQueue.drain();
    await emailQueue.obliterate({force:true});
    console.log("Server started");
});