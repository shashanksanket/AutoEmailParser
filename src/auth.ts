import passport from "passport";
import { Strategy as GoogleStrategy } from 'passport-google-oauth2';
import { google, Auth, gmail_v1 } from 'googleapis';
import { Client } from '@microsoft/microsoft-graph-client';
import { Strategy as MicrosoftStrategy } from "passport-microsoft";

import dotenv from "dotenv";
dotenv.config();

const getGraphClient = (accessToken: string): Client => {
    const client = Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });

    return client;
};

passport.use(
    new MicrosoftStrategy(
        {
            clientID: process.env.OUTLOOK_CLIENT_ID || '',
            clientSecret: process.env.OUTLOOK_CLIENT_SECRET || '',
            callbackURL: "http://localhost:5500/auth/outlook/callback",
        },
        async (accessToken: string, refreshToken: any, profile: any, done: (arg0: unknown, arg1: { id: any; displayName: any; email: any; accessToken: any; } | unknown) => any) => {
            try {
                const graphClient = getGraphClient(accessToken);

                const userProfile = await graphClient.api('/me').get();
                const user = {
                    id: userProfile.id,
                    displayName: userProfile.displayName,
                    email: userProfile.mail || userProfile.userPrincipalName,
                    accessToken: accessToken
                };

                return done(null, user);
            } catch (error) {
                console.error("Error during Outlook authentication:", error);
                return done(null,error);
            }
        }
    )
);

passport.use(
    new GoogleStrategy({
        clientID: process.env.GOOGLE_CLIENT_ID || '',
        clientSecret: process.env.GOOGLE_CLIENT_SECRET || '',
        callbackURL: "http://localhost:5500/auth/google/callback",
        passReqToCallback: true
    },
    (request: any, accessToken: any, refreshToken: any, profile: { tokens: { access_token: any; refresh_token: any; }; gmailProfile: gmail_v1.Schema$Profile; }, done: (arg0: Error | null, arg1: any) => void) => {
        profile.tokens = { access_token: accessToken, refresh_token: refreshToken };

        const oauth2Client = new google.auth.OAuth2();
        oauth2Client.setCredentials({ access_token: accessToken });
        const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
        gmail.users.getProfile({ userId: 'me' }, (err, res) => {
            if (err) {
                console.error('Error fetching Gmail profile:', err);
                return done(null,err);
            }
            profile.gmailProfile = res!.data;
            done(null, profile);
        });
    }
));

passport.serializeUser((user, done) => {
    done(null, user);
});

passport.deserializeUser((user, done) => {
    done(null, user);
});

export default passport;
