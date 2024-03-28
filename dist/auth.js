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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const passport_1 = __importDefault(require("passport"));
const passport_google_oauth2_1 = require("passport-google-oauth2");
const googleapis_1 = require("googleapis");
const microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
const passport_microsoft_1 = require("passport-microsoft");
const dotenv_1 = __importDefault(require("dotenv"));
dotenv_1.default.config();
// Function to initialize Microsoft Graph client
const getGraphClient = (accessToken) => {
    // Initialize Graph client with access token
    const client = microsoft_graph_client_1.Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
    return client;
};
passport_1.default.use(new passport_microsoft_1.Strategy({
    clientID: process.env.OUTLOOK_CLIENT_ID || '',
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET || '',
    callbackURL: "http://localhost:5500/auth/outlook/callback",
}, (accessToken, refreshToken, profile, done) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        const graphClient = getGraphClient(accessToken); // Implement a function to get Microsoft Graph client
        const userProfile = yield graphClient.api('/me').get();
        const user = {
            id: userProfile.id,
            displayName: userProfile.displayName,
            email: userProfile.mail || userProfile.userPrincipalName,
            accessToken: accessToken
        };
        return done(null, user);
    }
    catch (error) {
        console.error("Error during Outlook authentication:", error);
        return done(null, error);
    }
})));
passport_1.default.use(new passport_google_oauth2_1.Strategy({
    clientID: process.env.GOOGLE_CLIENT_ID || '',
    clientSecret: process.env.GOOGLE_CLIENT_SECRET || '',
    callbackURL: "http://localhost:5500/auth/google/callback",
    passReqToCallback: true
}, (request, accessToken, refreshToken, profile, done) => {
    profile.tokens = { access_token: accessToken, refresh_token: refreshToken };
    const oauth2Client = new googleapis_1.google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: accessToken });
    const gmail = googleapis_1.google.gmail({ version: 'v1', auth: oauth2Client });
    gmail.users.getProfile({ userId: 'me' }, (err, res) => {
        if (err) {
            console.error('Error fetching Gmail profile:', err);
            return done(null, err);
        }
        profile.gmailProfile = res.data;
        done(null, profile);
    });
}));
passport_1.default.serializeUser((user, done) => {
    done(null, user);
});
passport_1.default.deserializeUser((user, done) => {
    done(null, user);
});
exports.default = passport_1.default;
//# sourceMappingURL=auth.js.map