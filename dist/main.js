"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
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
const msal = __importStar(require("@azure/msal-node"));
const clientId = '8ddb0f11-add0-4145-87b5-604dd24216c2';
const clientSecret = '4Xw8Q~AsghPesFDFtbIPyJR6SK4bXhS8TazAwaYx';
const tenantId = 'f357f3b7-1fa3-4bc2-a3d8-22854f7e333a';
const scopes = ['https://graph.microsoft.com/.default'];
const authClient = new msal.ConfidentialClientApplication({
    auth: {
        clientId,
        clientSecret,
        authority: `https://login.microsoftonline.com/${tenantId}`,
    },
});
function getToken() {
    return authClient.acquireTokenByClientCredential({
        scopes,
    }).then(authResult => {
        if (!authResult) {
            throw new Error('authentication failed');
        }
        return authResult.accessToken;
    });
}
getToken()
    .then(accessToken => {
    console.log(`Access token: ${accessToken}`);
});
function sendEmail() {
    return __awaiter(this, void 0, void 0, function* () {
        const accessToken = yield getToken();
        const mailuser = '360Reviews.IOCO@eoh.com';
        const Response = yield fetch(`https://graph.microsoft.com/v1.0/users/${mailuser}/sendMail` /**https://graph.microsoft.com/v1.0/me/sendMailview word wrap */, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                message: {
                    subject: 'Test email from Typescript',
                    toRecipients: [
                        {
                            emailAddress: {
                                address: 'thembelani.mkhize@eoh.com',
                            },
                        },
                    ],
                    body: {
                        content: 'This is the test content body from typescript',
                        contentType: 'Text',
                    },
                },
            }),
        });
        if (!Response.ok) {
            const text = yield Response.text();
            throw new Error(`Failed to send email: ${text}`);
        }
    });
}
sendEmail()
    .then(() => {
    console.log('The email was send successfully');
})
    .catch(error => {
    console.log(`Error sending the email: ${error}`);
});
// const options = {
// 	authProvider,
// };
// const client = Client.init(options);
// const sendMail = {message: {subject: 'Meet for lunch?',body: {contentType: 'Text',content: 'The new cafeteria is open.'},toRecipients: [{emailAddress: {address: 'garthf@contoso.com'}}]}};
// await client.api('/me/sendMail')
// 	.post(sendMail); node dist/main.js
// Code snippets are only available for the latest version. Current version is 6.x
// GraphServiceClient graphClient = new GraphServiceClient(requestAdapter);
// com.microsoft.graph.users.item.sendmail.SendMailPostRequestBody sendMailPostRequestBody = new com.microsoft.graph.users.item.sendmail.SendMailPostRequestBody();
// Message message = new Message();
// message.setSubject("Meet for lunch?");
// ItemBody body = new ItemBody();
// body.setContentType(BodyType.Text);
// body.setContent("The new cafeteria is open.");
// message.setBody(body);
// LinkedList<Recipient> toRecipients = new LinkedList<Recipient>();
// Recipient recipient = new Recipient();
// EmailAddress emailAddress = new EmailAddress();
// emailAddress.setAddress("thembelani.mkhize@eoh.com");
// recipient.setEmailAddress(emailAddress);
// toRecipients.add(recipient);
// message.setToRecipients(toRecipients);
// sendMailPostRequestBody.setMessage(message);
// graphClient.me().sendMail().post(sendMailPostRequestBody);
