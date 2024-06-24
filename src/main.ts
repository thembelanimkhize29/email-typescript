import * as msal from '@azure/msal-node';
import { RequestInit, Response } from 'node-fetch';

const clientId='8ddb0f11-add0-4145-87b5-604dd24216c2';
const clientSecret='4Xw8Q~AsghPesFDFtbIPyJR6SK4bXhS8TazAwaYx';
const tenantId='f357f3b7-1fa3-4bc2-a3d8-22854f7e333a';
const scopes=['https://graph.microsoft.com/.default'];

const authClient = new msal.ConfidentialClientApplication({
    auth:{
        clientId,
        clientSecret,
        authority: `https://login.microsoftonline.com/${tenantId}`,
    },

});

function getToken(): Promise<string>{
    return authClient.acquireTokenByClientCredential({
        scopes,
    }).then(authResult => {
        if(!authResult){
            throw new Error('authentication failed');
        }

        return authResult.accessToken;
    });
}

getToken()
.then(accessToken => {
    console.log(`Access token: ${accessToken}`);
    
});

async function sendEmail(): Promise<void> {
    const accessToken =await getToken();
    const mailuser='360Reviews.IOCO@eoh.com';
    const Response = await fetch(`https://graph.microsoft.com/v1.0/users/${mailuser}/sendMail`/**https://graph.microsoft.com/v1.0/me/sendMailview word wrap */,{
        method: 'POST',
        headers:{
            'Authorization':`Bearer ${accessToken}`,
            'Content-Type':'application/json',
        },
        body: JSON.stringify({
            message:{
                subject: 'Test email from Typescript',
                toRecipients:[
                    {
                        emailAddress :{
                            address:'thembelani.mkhize@eoh.com',
                        },
                    },
                ],
                body :{
                    content: 'This is the test content body from typescript',
                    contentType:'Text',
                },
            },
        }),

    });

    if(!Response.ok){
        const text=await Response.text();
        throw new Error(`Failed to send email: ${text}`);
    }
}

sendEmail()
.then(()=>{
    console.log('The email was send successfully');
})
.catch(error=>{
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