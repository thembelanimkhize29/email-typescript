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
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var msal = require("@azure/msal-node");
var clientId = '8ddb0f11-add0-4145-87b5-604dd24216c2';
var clientSecret = '4Xw8Q~AsghPesFDFtbIPyJR6SK4bXhS8TazAwaYx';
var tenantId = 'f357f3b7-1fa3-4bc2-a3d8-22854f7e333a';
var scopes = ['https://graph.microsoft.com/.default'];
var authClient = new msal.ConfidentialClientApplication({
    auth: {
        clientId: clientId,
        clientSecret: clientSecret,
        authority: "https://login.microsoftonline.com/".concat(tenantId),
    },
});
function getToken() {
    return authClient.acquireTokenByClientCredential({
        scopes: scopes,
    }).then(function (authResult) {
        if (!authResult) {
            throw new Error('authentication failed');
        }
        return authResult.accessToken;
    });
}
getToken()
    .then(function (accessToken) {
    console.log("Access token: ".concat(accessToken));
});
function sendEmail() {
    return __awaiter(this, void 0, void 0, function () {
        var accessToken, mailuser, Response, text;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getToken()];
                case 1:
                    accessToken = _a.sent();
                    mailuser = 'engage@eoh.com';
                    return [4 /*yield*/, fetch("https://graph.microsoft.com/v1.0/users/".concat(mailuser, "/sendMail") /**view word wrap */, {
                            method: 'POST',
                            headers: {
                                'Authorization': "Bearer ".concat(accessToken),
                                'Content-Type': 'application/json',
                            },
                            body: JSON.stringify({
                                message: {
                                    subject: 'Test email from Typescript',
                                    toRecipients: [
                                        {
                                            emailAddress: {
                                                address: 'thembelanimkhize29@gmail.com',
                                            },
                                        },
                                    ],
                                    body: {
                                        content: 'This is the test content body from typescript',
                                        contentType: 'Text',
                                    },
                                },
                            }),
                        })];
                case 2:
                    Response = _a.sent();
                    if (!!Response.ok) return [3 /*break*/, 4];
                    return [4 /*yield*/, Response.text()];
                case 3:
                    text = _a.sent();
                    throw new Error("1 Failed to send email: ".concat(text));
                case 4: return [2 /*return*/];
            }
        });
    });
}
sendEmail()
    .then(function () {
    console.log('The email was send successfully');
})
    .catch(function (error) {
    console.log("Error sending the email: ".concat(error));
});
