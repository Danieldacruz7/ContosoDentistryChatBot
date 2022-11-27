// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler } = require('botbuilder');
var request = require('request');
const dotenv = require('dotenv');
const axios = require('axios');

const ENV_FILE = './.env';
dotenv.config({ path: ENV_FILE });

/**
 * A simple bot that responds to utterances with answers from QnA Maker.
 * If an answer is not found for an utterance, the bot responds with help.
 */
class QnABot extends ActivityHandler {
    /**
     *
     * @param {ConversationState} conversationState
     * @param {UserState} userState
     * @param {Dialog} dialog
     */
    constructor(conversationState, userState, dialog) {
        super();
        if (!conversationState) throw new Error('[QnABot]: Missing parameter. conversationState is required');
        if (!userState) throw new Error('[QnABot]: Missing parameter. userState is required');
        if (!dialog) throw new Error('[QnABot]: Missing parameter. dialog is required');

        this.conversationState = conversationState;
        this.userState = userState;
        this.dialog = dialog;
        this.dialogState = this.conversationState.createProperty('DialogState');

        this.onMessage(async (context, next) => {
            console.log('Running dialog with Message Activity.');

            const response = await axios.post(
                'https://dental-assistant-langservice.cognitiveservices.azure.com/language/:analyze-conversations',
                {
                    'kind': 'Conversation',
                    'analysisInput': {
                        'conversationItem': {
                            'id': '1',
                            'text': context._activity.text,
                            'modality': 'text',
                            'language': 'en-US',
                            'participantId': '1'
                        }
                    },
                    'parameters': {
                        'projectName': 'dental-assistant-workflow',
                        'verbose': true,
                        'deploymentName': 'dental-assistant-deployment',
                        'stringIndexType': 'TextElement_V8'
                    }
                },
                {
                    params: {
                        'api-version': '2022-10-01-preview'
                    },
                    headers: {
                        'Ocp-Apim-Subscription-Key': 'a0485a45bcf0479290891c881dcf5b01',
                        'Apim-Request-Id': '4ffcac1c-b2fc-48ba-bd6d-b69d9942995a',
                        'Content-Type': 'application/json'
                    }
                }
            );

            if (response.status === 200){
            if (response['data']['result']['prediction']['topIntent'] === 'GetAvailability' 
                && response['data']['result']['prediction']['intents']['GetAvailability']['confidenceScore'] > 0.5){
                    const answer = await axios.get("https://dental-assistant-scheduler10.azurewebsites.net/availability");
                    await context.sendActivity(`Here is a list of all the available times: ${answer['data']}`);
                        }

           if (response['data']['result']['prediction']['topIntent'] === 'ScheduleAppointment' 
                && response['data']['result']['prediction']['intents']['ScheduleAppointment']['confidenceScore'] > 0.5){
                    const schedule = await axios.post(
                        "https://dental-assistant-scheduler10.azurewebsites.net/schedule",
                        {
                            'Time': response['data']['result']['prediction']['intents']['ScheduleAppointment']['result']['prediction']['entities'][0]['text']
                        },
                        {
                            headers: {
                                'Content-Type': 'application/json'
                            }
                        }
                    );
                    await context.sendActivity(`Your appointment has been set for ${response['data']['result']['prediction']['intents']['ScheduleAppointment']['result']['prediction']['entities'][0]['text']}`);
                    }

            if (response['data']['result']['prediction']['topIntent'] === 'GeneralQueries'
            && response['data']['result']['prediction']['intents']['GeneralQueries']['confidenceScore'] > 0.5){
                           await context.sendActivity(response['data']['result']['prediction']['intents']['GeneralQueries']['result']['answers'][0]['answer']);
                        }
            if (response['data']['result']['prediction']['topIntent'] === 'ScheduleAppointment' 
                && response['data']['result']['prediction']['intents']['ScheduleAppointment']['confidenceScore'] < 0.5
                ||response['data']['result']['prediction']['topIntent'] === 'GeneralQueries' 
                && response['data']['result']['prediction']['intents']['GeneralQueries']['confidenceScore'] < 0.5
                ||response['data']['result']['prediction']['topIntent'] === 'GetAvailability' 
                && response['data']['result']['prediction']['intents']['GetAvailability']['confidenceScore'] < 0.5){
                    await context.sendActivity("I didn't quite understand. Please try again...");
                }
            }
            else{context.sendActivity("Please try again.");};


            // Run the Dialog with the new message Activity.
            //await this.dialog.run(context, this.dialogState);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });


        // If a new user is added to the conversation, send them a greeting message
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    const defaultWelcome = process.env.DefaultWelcomeMessage;
                    if (defaultWelcome !== '') await context.sendActivity(defaultWelcome);
                    else await context.sendActivity('Welcome to Contoso Dentistry! How can I help?');
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    /**
     * Override the ActivityHandler.run() method to save state changes after the bot logic completes.
     */
    async run(context) {
        await super.run(context);

        // Save any state changes. The load happened during the execution of the Dialog.
        await this.conversationState.saveChanges(context, false);
        await this.userState.saveChanges(context, false);
    }
}

module.exports.QnABot = QnABot;

// SIG // Begin signature block
// SIG // MIIrYgYJKoZIhvcNAQcCoIIrUzCCK08CAQExDzANBglg
// SIG // hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
// SIG // BgEEAYI3AgEeMCQCAQEEEBDgyQbOONQRoqMAEEvTUJAC
// SIG // AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
// SIG // 0eQxtlB2VyiNr0ZioFaUYLkhQ39TDEmM1GEvfBIe1bGg
// SIG // ghF5MIIIiTCCB3GgAwIBAgITNgAAAanWkDBmQ9sfggAC
// SIG // AAABqTANBgkqhkiG9w0BAQsFADBBMRMwEQYKCZImiZPy
// SIG // LGQBGRYDR0JMMRMwEQYKCZImiZPyLGQBGRYDQU1FMRUw
// SIG // EwYDVQQDEwxBTUUgQ1MgQ0EgMDEwHhcNMjIwNjEwMTgy
// SIG // NzA0WhcNMjMwNjEwMTgyNzA0WjAkMSIwIAYDVQQDExlN
// SIG // aWNyb3NvZnQgQXp1cmUgQ29kZSBTaWduMIIBIjANBgkq
// SIG // hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAuLvS3Hq6XM6N
// SIG // 5ZVPdqZQQbTo4WFo9Ar6TqyLpZIqQpNoW9ZG58deayDX
// SIG // VV7wKgn0IAjewM3VfPGtiX8jjOz4VtelbCYnbV9zrqqU
// SIG // rtTlqTbFB1L+UWQO2DLhxB8QybLxi38KaiY1DC6DL5xK
// SIG // uAnIGWnVNS168FihSxIPneGKfG3nJH1CgSA/rJ7w7tnY
// SIG // 8hHlpPpMia6oKVAZSvos9/fDpBmX+cru3sXfEv19s+4O
// SIG // JKLoPlJiNR0PhsqW5hChTn+tjVOBu8Td7tcb+jf9QQs1
// SIG // 2HPBtx3nMNhNlYZQrqXJMUy65RH2zAYAd9N9tdo6VRU/
// SIG // 8vRYzYOrWHSulDVtMn2cjwIDAQABo4IFlTCCBZEwKQYJ
// SIG // KwYBBAGCNxUKBBwwGjAMBgorBgEEAYI3WwEBMAoGCCsG
// SIG // AQUFBwMDMD0GCSsGAQQBgjcVBwQwMC4GJisGAQQBgjcV
// SIG // CIaQ4w2E1bR4hPGLPoWb3RbOnRKBYIPdzWaGlIwyAgFk
// SIG // AgEMMIICdgYIKwYBBQUHAQEEggJoMIICZDBiBggrBgEF
// SIG // BQcwAoZWaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3Br
// SIG // aWluZnJhL0NlcnRzL0JZMlBLSUNTQ0EwMS5BTUUuR0JM
// SIG // X0FNRSUyMENTJTIwQ0ElMjAwMSgyKS5jcnQwUgYIKwYB
// SIG // BQUHMAKGRmh0dHA6Ly9jcmwxLmFtZS5nYmwvYWlhL0JZ
// SIG // MlBLSUNTQ0EwMS5BTUUuR0JMX0FNRSUyMENTJTIwQ0El
// SIG // MjAwMSgyKS5jcnQwUgYIKwYBBQUHMAKGRmh0dHA6Ly9j
// SIG // cmwyLmFtZS5nYmwvYWlhL0JZMlBLSUNTQ0EwMS5BTUUu
// SIG // R0JMX0FNRSUyMENTJTIwQ0ElMjAwMSgyKS5jcnQwUgYI
// SIG // KwYBBQUHMAKGRmh0dHA6Ly9jcmwzLmFtZS5nYmwvYWlh
// SIG // L0JZMlBLSUNTQ0EwMS5BTUUuR0JMX0FNRSUyMENTJTIw
// SIG // Q0ElMjAwMSgyKS5jcnQwUgYIKwYBBQUHMAKGRmh0dHA6
// SIG // Ly9jcmw0LmFtZS5nYmwvYWlhL0JZMlBLSUNTQ0EwMS5B
// SIG // TUUuR0JMX0FNRSUyMENTJTIwQ0ElMjAwMSgyKS5jcnQw
// SIG // ga0GCCsGAQUFBzAChoGgbGRhcDovLy9DTj1BTUUlMjBD
// SIG // UyUyMENBJTIwMDEsQ049QUlBLENOPVB1YmxpYyUyMEtl
// SIG // eSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZp
// SIG // Z3VyYXRpb24sREM9QU1FLERDPUdCTD9jQUNlcnRpZmlj
// SIG // YXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlmaWNhdGlv
// SIG // bkF1dGhvcml0eTAdBgNVHQ4EFgQUj5gJWFiDzm06yLnX
// SIG // Wf2V9PM6+1cwDgYDVR0PAQH/BAQDAgeAMFAGA1UdEQRJ
// SIG // MEekRTBDMSkwJwYDVQQLEyBNaWNyb3NvZnQgT3BlcmF0
// SIG // aW9ucyBQdWVydG8gUmljbzEWMBQGA1UEBRMNMjM2MTY3
// SIG // KzQ3MDg2MTCCAeYGA1UdHwSCAd0wggHZMIIB1aCCAdGg
// SIG // ggHNhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
// SIG // aW5mcmEvQ1JML0FNRSUyMENTJTIwQ0ElMjAwMSgyKS5j
// SIG // cmyGMWh0dHA6Ly9jcmwxLmFtZS5nYmwvY3JsL0FNRSUy
// SIG // MENTJTIwQ0ElMjAwMSgyKS5jcmyGMWh0dHA6Ly9jcmwy
// SIG // LmFtZS5nYmwvY3JsL0FNRSUyMENTJTIwQ0ElMjAwMSgy
// SIG // KS5jcmyGMWh0dHA6Ly9jcmwzLmFtZS5nYmwvY3JsL0FN
// SIG // RSUyMENTJTIwQ0ElMjAwMSgyKS5jcmyGMWh0dHA6Ly9j
// SIG // cmw0LmFtZS5nYmwvY3JsL0FNRSUyMENTJTIwQ0ElMjAw
// SIG // MSgyKS5jcmyGgb1sZGFwOi8vL0NOPUFNRSUyMENTJTIw
// SIG // Q0ElMjAwMSgyKSxDTj1CWTJQS0lDU0NBMDEsQ049Q0RQ
// SIG // LENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNl
// SIG // cnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9QU1FLERD
// SIG // PUdCTD9jZXJ0aWZpY2F0ZVJldm9jYXRpb25MaXN0P2Jh
// SIG // c2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9p
// SIG // bnQwHwYDVR0jBBgwFoAUllGE4Gtve/7YBqvD8oXmKa5q
// SIG // +dQwHwYDVR0lBBgwFgYKKwYBBAGCN1sBAQYIKwYBBQUH
// SIG // AwMwDQYJKoZIhvcNAQELBQADggEBAHD1OJbFZ/tIa5Zp
// SIG // DzeU+mqWHOdF2htAZKicRfNYhaajjyYRvCTUKn/5SZGU
// SIG // KKdVmsxiFtCOp2lJ2+C3b7IJukkqC9SmpIkQLhBuz7uK
// SIG // 4NsXB6Xn3Iv32YuKeH4sqdRqJMCezhsale/Sh6fecsVW
// SIG // pJnsvfXxdXBCyoVbAZCZCQN3dOXUz4DtEfV2fxhRzTfS
// SIG // UhKsr1VSY9HC/myediSqvqd3zfgK9j6IR0DcL3WkKiV0
// SIG // B/dnYwntnntrhFxGYQuPPXBA7xX10SB/8CVA8V1NovOk
// SIG // tGO5cgvmVMe5pA2m9M7sOBgFkjXgPD7i4PoL5X0mK+6b
// SIG // nchiEZj1C5l1X6LzJH4wggjoMIIG0KADAgECAhMfAAAA
// SIG // UeqP9pxzDKg7AAAAAABRMA0GCSqGSIb3DQEBCwUAMDwx
// SIG // EzARBgoJkiaJk/IsZAEZFgNHQkwxEzARBgoJkiaJk/Is
// SIG // ZAEZFgNBTUUxEDAOBgNVBAMTB2FtZXJvb3QwHhcNMjEw
// SIG // NTIxMTg0NDE0WhcNMjYwNTIxMTg1NDE0WjBBMRMwEQYK
// SIG // CZImiZPyLGQBGRYDR0JMMRMwEQYKCZImiZPyLGQBGRYD
// SIG // QU1FMRUwEwYDVQQDEwxBTUUgQ1MgQ0EgMDEwggEiMA0G
// SIG // CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDJmlIJfQGe
// SIG // jVbXKpcyFPoFSUllalrinfEV6JMc7i+bZDoL9rNHnHDG
// SIG // fJgeuRIYO1LY/1f4oMTrhXbSaYRCS5vGc8145WcTZG90
// SIG // 8bGDCWr4GFLc411WxA+Pv2rteAcz0eHMH36qTQ8L0o3X
// SIG // Ob2n+x7KJFLokXV1s6pF/WlSXsUBXGaCIIWBXyEchv+s
// SIG // M9eKDsUOLdLTITHYJQNWkiryMSEbxqdQUTVZjEz6eLRL
// SIG // kofDAo8pXirIYOgM770CYOiZrcKHK7lYOVblx22pdNaw
// SIG // Y8Te6a2dfoCaWV1QUuazg5VHiC4p/6fksgEILptOKhx9
// SIG // c+iapiNhMrHsAYx9pUtppeaFAgMBAAGjggTcMIIE2DAS
// SIG // BgkrBgEEAYI3FQEEBQIDAgACMCMGCSsGAQQBgjcVAgQW
// SIG // BBQSaCRCIUfL1Gu+Mc8gpMALI38/RzAdBgNVHQ4EFgQU
// SIG // llGE4Gtve/7YBqvD8oXmKa5q+dQwggEEBgNVHSUEgfww
// SIG // gfkGBysGAQUCAwUGCCsGAQUFBwMBBggrBgEFBQcDAgYK
// SIG // KwYBBAGCNxQCAQYJKwYBBAGCNxUGBgorBgEEAYI3CgMM
// SIG // BgkrBgEEAYI3FQYGCCsGAQUFBwMJBggrBgEFBQgCAgYK
// SIG // KwYBBAGCN0ABAQYLKwYBBAGCNwoDBAEGCisGAQQBgjcK
// SIG // AwQGCSsGAQQBgjcVBQYKKwYBBAGCNxQCAgYKKwYBBAGC
// SIG // NxQCAwYIKwYBBQUHAwMGCisGAQQBgjdbAQEGCisGAQQB
// SIG // gjdbAgEGCisGAQQBgjdbAwEGCisGAQQBgjdbBQEGCisG
// SIG // AQQBgjdbBAEGCisGAQQBgjdbBAIwGQYJKwYBBAGCNxQC
// SIG // BAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMBIGA1Ud
// SIG // EwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAUKV5RXmSu
// SIG // NLnrrJwNp4x1AdEJCygwggFoBgNVHR8EggFfMIIBWzCC
// SIG // AVegggFToIIBT4YxaHR0cDovL2NybC5taWNyb3NvZnQu
// SIG // Y29tL3BraWluZnJhL2NybC9hbWVyb290LmNybIYjaHR0
// SIG // cDovL2NybDIuYW1lLmdibC9jcmwvYW1lcm9vdC5jcmyG
// SIG // I2h0dHA6Ly9jcmwzLmFtZS5nYmwvY3JsL2FtZXJvb3Qu
// SIG // Y3JshiNodHRwOi8vY3JsMS5hbWUuZ2JsL2NybC9hbWVy
// SIG // b290LmNybIaBqmxkYXA6Ly8vQ049YW1lcm9vdCxDTj1B
// SIG // TUVSb290LENOPUNEUCxDTj1QdWJsaWMlMjBLZXklMjBT
// SIG // ZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0
// SIG // aW9uLERDPUFNRSxEQz1HQkw/Y2VydGlmaWNhdGVSZXZv
// SIG // Y2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNzPWNSTERp
// SIG // c3RyaWJ1dGlvblBvaW50MIIBqwYIKwYBBQUHAQEEggGd
// SIG // MIIBmTBHBggrBgEFBQcwAoY7aHR0cDovL2NybC5taWNy
// SIG // b3NvZnQuY29tL3BraWluZnJhL2NlcnRzL0FNRVJvb3Rf
// SIG // YW1lcm9vdC5jcnQwNwYIKwYBBQUHMAKGK2h0dHA6Ly9j
// SIG // cmwyLmFtZS5nYmwvYWlhL0FNRVJvb3RfYW1lcm9vdC5j
// SIG // cnQwNwYIKwYBBQUHMAKGK2h0dHA6Ly9jcmwzLmFtZS5n
// SIG // YmwvYWlhL0FNRVJvb3RfYW1lcm9vdC5jcnQwNwYIKwYB
// SIG // BQUHMAKGK2h0dHA6Ly9jcmwxLmFtZS5nYmwvYWlhL0FN
// SIG // RVJvb3RfYW1lcm9vdC5jcnQwgaIGCCsGAQUFBzAChoGV
// SIG // bGRhcDovLy9DTj1hbWVyb290LENOPUFJQSxDTj1QdWJs
// SIG // aWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxD
// SIG // Tj1Db25maWd1cmF0aW9uLERDPUFNRSxEQz1HQkw/Y0FD
// SIG // ZXJ0aWZpY2F0ZT9iYXNlP29iamVjdENsYXNzPWNlcnRp
// SIG // ZmljYXRpb25BdXRob3JpdHkwDQYJKoZIhvcNAQELBQAD
// SIG // ggIBAFAQI7dPD+jfXtGt3vJp2pyzA/HUu8hjKaRpM3op
// SIG // ya5G3ocprRd7vdTHb8BDfRN+AD0YEmeDB5HKQoG6xHPI
// SIG // 5TXuIi5sm/LeADbV3C2q0HQOygS/VT+m1W7a/752hMIn
// SIG // +L4ZuyxVeSBpfwf7oQ4YSZPh6+ngZvBHgfBaVz4O9/wc
// SIG // fw91QDZnTgK9zAh9yRKKls2bziPEnxeOZMVNaxyV0v15
// SIG // 2PY2xjqIafIkUjK6vY9LtVFjJXenVUAmn3WCPWNFC1YT
// SIG // IIHw/mD2cTfPy7QA1pT+GPARAKt0bKtq9aCd/Ym0b5tP
// SIG // bpgCiRtzyb7fbNS1dE740re0COE67YV2wbeo2sXixzvL
// SIG // ftH8L7s9xv9wV+G22qyKt6lmKLjFK1yMw4Ni5fMabcgm
// SIG // zRvSjAcbqgp3tk4a8emaaH0rz8MuuIP+yrxtREPXSqL/
// SIG // C5bzMzsikuDW9xH10graZzSmPjilzpRfRdu20/9UQmC7
// SIG // eVPZ4j1WNa1oqPHfzET3ChIzJ6Q9G3NPCB+7KwX0OQmK
// SIG // yv7IDimj8U/GlsHD1z+EF/fYMf8YXG15LamaOAohsw/y
// SIG // wO6SYSreVW+5Y0mzJutnBC9Cm9ozj1+/4kqksrlhZgR/
// SIG // CSxhFH3BTweH8gP2FEISRtShDZbuYymynY1un+RyfiK9
// SIG // +iVTLdD1h/SxyxDpZMtimb4CgJQlMYIZQTCCGT0CAQEw
// SIG // WDBBMRMwEQYKCZImiZPyLGQBGRYDR0JMMRMwEQYKCZIm
// SIG // iZPyLGQBGRYDQU1FMRUwEwYDVQQDEwxBTUUgQ1MgQ0Eg
// SIG // MDECEzYAAAGp1pAwZkPbH4IAAgAAAakwDQYJYIZIAWUD
// SIG // BAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
// SIG // AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
// SIG // LwYJKoZIhvcNAQkEMSIEIEE8yjII+wdKaPwu8JBGRHPh
// SIG // WLJIZ7Fqp4F3A143c9KpMEIGCisGAQQBgjcCAQwxNDAy
// SIG // oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8v
// SIG // d3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEBBQAE
// SIG // ggEARpnbwXsqVO60lqJPqiWOjOI03i77Y0pYXQXlu2gZ
// SIG // fVdmEKVVzXc6YUoj52wy3M2LEjLYiHnigaEEigN6Us80
// SIG // BqyVdP8VaLuD4eBzoDXzR3uhbGMe9w2DigeK0HHQSe5z
// SIG // sb+GyhUoPlCTLvuFKzQ1nMls2pfrmvh+ItzLJ8s7l9u3
// SIG // kdB21dEE6taprn8mA9RYcODb8luQCruLIlRAjZMeJg8S
// SIG // j922Rx4Cz9kw6ex6eLR/itk49kGJPmhR1bm92z8u3KgR
// SIG // zZzdxYuh1KU7KaYDwVqjJblS03E4HcFi1YqcmGygWcMm
// SIG // Nqn+k7tIbZLqGEgcpdyJ7ESJ/NJeYVKxwrBFuKGCFwkw
// SIG // ghcFBgorBgEEAYI3AwMBMYIW9TCCFvEGCSqGSIb3DQEH
// SIG // AqCCFuIwghbeAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFV
// SIG // BgsqhkiG9w0BCRABBKCCAUQEggFAMIIBPAIBAQYKKwYB
// SIG // BAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCBJKnLBW8GN
// SIG // KACZV5MxqLuxpkvrbNmESGSnDC+OOpfIbgIGY2K8qib9
// SIG // GBMyMDIyMTEwMzA4MDU1Mi44MzhaMASAAgH0oIHUpIHR
// SIG // MIHOMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
// SIG // Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
// SIG // TWljcm9zb2Z0IENvcnBvcmF0aW9uMSkwJwYDVQQLEyBN
// SIG // aWNyb3NvZnQgT3BlcmF0aW9ucyBQdWVydG8gUmljbzEm
// SIG // MCQGA1UECxMdVGhhbGVzIFRTUyBFU046RjdBNi1FMjUx
// SIG // LTE1MEExJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0
// SIG // YW1wIFNlcnZpY2WgghFcMIIHEDCCBPigAwIBAgITMwAA
// SIG // AaUA3gjEQAdxTgABAAABpTANBgkqhkiG9w0BAQsFADB8
// SIG // MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
// SIG // bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
// SIG // cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
// SIG // b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMjAz
// SIG // MDIxODUxMTlaFw0yMzA1MTExODUxMTlaMIHOMQswCQYD
// SIG // VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
// SIG // A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
// SIG // IENvcnBvcmF0aW9uMSkwJwYDVQQLEyBNaWNyb3NvZnQg
// SIG // T3BlcmF0aW9ucyBQdWVydG8gUmljbzEmMCQGA1UECxMd
// SIG // VGhhbGVzIFRTUyBFU046RjdBNi1FMjUxLTE1MEExJTAj
// SIG // BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZp
// SIG // Y2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
// SIG // AQC6sYboIGpIvMLqDjDHe67BEJ5gIbVfIlNWNIrbB6t9
// SIG // E3QlyQ5r2Y2mfMrzh2BVYU8g9W+SRibcGY1s9X4JQqrM
// SIG // eagcT9VsdQmZ7ENbYkbEVkHNdlZBE5pGPMeOjIB7BsgJ
// SIG // oTz6bIEZ5JRmoux6kBQd9cf0I5Me62wJa+j25QeLTpmk
// SIG // dZysZeFSILLQ8H53imqBBMOIjf8U3c7WY8MhomOYTaem
// SIG // 3nrZHIs4CRTt/8kR2IdILZPm0RIa5iIG2q664G8+zLJw
// SIG // O7ZSrxnDvYh3OvtrMpqwFctws0OCDDTxXE08fME2fpKb
// SIG // +pRbNXhvMZX7LtjQ1irIazJSh9iaWM1gFtXwjg+Yq17B
// SIG // OCzr4sWUL253kBOvohnyEMGm4/n0XaLgFNgIhPomjbCA
// SIG // 2qXSmm/Fi8c+lT0WxC/jOjBZHLKIrihx6LIQqeyYZmfY
// SIG // jNMqxMdl3mzoWv10N+NirERrNodNoKV+sAcsk/Hg9zCV
// SIG // SMUkZuDCyIpb1nKXfTd66KGsGy1OoHZO4KClkuvfsNo7
// SIG // aLlwhGLeiD32avJXYtC/wsGG7b+5mx5iGfTnNCRCXOm/
// SIG // YHFQ36D4npjCnM9eQS3qcse56UNjIgyiLHDqioV7mSPj
// SIG // 2XqzTh4Yv77MtvxY/ZQepCazGEn1dBdn67wUgVzAe8Y7
// SIG // /KYKl+UF1HvJ08W+FHydHAwLwQIDAQABo4IBNjCCATIw
// SIG // HQYDVR0OBBYEFF+mjwMAl66urXDu+9xZF0toqRrfMB8G
// SIG // A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8G
// SIG // A1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9z
// SIG // b2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUyMFRp
// SIG // bWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggr
// SIG // BgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93
// SIG // d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWlj
// SIG // cm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAo
// SIG // MSkuY3J0MAwGA1UdEwEB/wQCMAAwEwYDVR0lBAwwCgYI
// SIG // KwYBBQUHAwgwDQYJKoZIhvcNAQELBQADggIBAJabCxfl
// SIG // MDCihEdqdFiZ6OBuhhhp34N6ow3Wh3Obr12LRuiph66g
// SIG // H/2Kh5JjaLUq+mRBJ5RgiWEe1t7ifuW6b49N8Bahnn70
// SIG // LCiEdvquk686M7z+DbKHVk0+UlafwukxAxriwvZjkCgO
// SIG // Lci+NB01u7cW9HAHX4J8hxaCPwbGaPxWl3s0PITuMVI4
// SIG // Q6cjTXielmL1+TQvh7/Z5k8s46shIPy9nFwDpsRFr3zw
// SIG // ENZX8b67VMBu+YxnlGnsJIcLc2pwpz95emI8CRSgep+/
// SIG // 017a34pNcWNZIHr9ScEOWlHT8cEnQ5hhOF0zdrOqTzov
// SIG // CDtffTn+gBL4eNXg8Uc/tdVVHKbhp+7SVHkk1Eh7L80P
// SIG // BAjo+cO+zL+efxfIVrtO3oJxvEq1o+fkxcTTwqcfwBTb
// SIG // 88/qHU0U2XeC1rqJnDB1JixYlBjgHXrRekqHxxuRHBZ9
// SIG // A0w9WqQWcwj/MbBkHGYMFaqO6L9t/7iCZTAiwMk2GVfS
// SIG // Ewj9PXIlCWygVQkDaxhJ0P1yxTvZsrMsg0a7x4VObhj3
// SIG // V8+Cbdv2TeyUGEblTUrgqTcKCtCa9bOnIg7xxHi8onM8
// SIG // aCHvRh90sn2x8er/6YSPohNw1qNUwiu+RC+qbepOYt+v
// SIG // 5J9rklV3Ux+OGVZId/4oVd7xMLO/Lhpb7IjHKygYKaNx
// SIG // 3XIwx4h6FrFH+BiMMIIHcTCCBVmgAwIBAgITMwAAABXF
// SIG // 52ueAptJmQAAAAAAFTANBgkqhkiG9w0BAQsFADCBiDEL
// SIG // MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
// SIG // EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
// SIG // c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9z
// SIG // b2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIw
// SIG // MTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgzMjI1
// SIG // WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
// SIG // Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
// SIG // TWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1N
// SIG // aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIw
// SIG // DQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAOThpkzn
// SIG // tHIhC3miy9ckeb0O1YLT/e6cBwfSqWxOdcjKNVf2AX9s
// SIG // SuDivbk+F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWN
// SIG // E893MsAQGOhgfWpSg0S3po5GawcU88V29YZQ3MFEyHFc
// SIG // UTE3oAo4bo3t1w/YJlN8OWECesSq/XJprx2rrPY2vjUm
// SIG // ZNqYO7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W
// SIG // 7IVWTe/dvI2k45GPsjksUZzpcGkNyjYtcI4xyDUoveO0
// SIG // hyTD4MmPfrVUj9z6BVWYbWg7mka97aSueik3rMvrg0Xn
// SIG // Rm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De
// SIG // +JKRHh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9fvzZnkXf
// SIG // tnIv231fgLrbqn427DZM9ituqBJR6L8FA6PRc6ZNN3SU
// SIG // HDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEzOUyO
// SIG // ArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC4jMYcten
// SIG // IPDC+hIK12NvDMk2ZItboKaDIV1fMHSRlJTYuVD5C4lh
// SIG // 8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqvUAV6bMUR
// SIG // HXLvjflSxIUXk8A8FdsaN8cIFRg/eKtFtvUeh17aj54W
// SIG // cmnGrnu3tz5q4i6tAgMBAAGjggHdMIIB2TASBgkrBgEE
// SIG // AYI3FQEEBQIDAQABMCMGCSsGAQQBgjcVAgQWBBQqp1L+
// SIG // ZMSavoKRPEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBe
// SIG // Yl2D9OXSZacbUzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEE
// SIG // AYI3TIN9AQEwQTA/BggrBgEFBQcCARYzaHR0cDovL3d3
// SIG // dy5taWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9z
// SIG // aXRvcnkuaHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBkG
// SIG // CSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQE
// SIG // AwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaA
// SIG // FNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRPME0w
// SIG // S6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9w
// SIG // a2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEw
// SIG // LTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYB
// SIG // BQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9w
// SIG // a2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMu
// SIG // Y3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9n
// SIG // ATEkW+Geckv8qW/qXBS2Pk5HZHixBpOXPTEztTnXwnE2
// SIG // P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7pZmc6U03dmLq2
// SIG // HnjYNi6cqYJWAAOwBb6J6Gngugnue99qb74py27YP0h1
// SIG // AdkY3m2CDPVtI1TkeFN1JFe53Z/zjj3G82jfZfakVqr3
// SIG // lbYoVSfQJL1AoL8ZthISEV09J+BAljis9/kpicO8F7BU
// SIG // hUKz/AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tu
// SIG // PywJeBTpkbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6MhrZ
// SIG // lvSP9pEB9s7GdP32THJvEKt1MMU0sHrYUP4KWN1APMdU
// SIG // bZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz
// SIG // 138eW0QBjloZkWsNn6Qo3GcZKCS6OEuabvshVGtqRRFH
// SIG // qfG3rsjoiV5PndLQTHa1V1QJsWkBRH58oWFsc/4Ku+xB
// SIG // Zj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLBgqJ7
// SIG // Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgkNWcr4A24
// SIG // 5oyZ1uEi6vAnQj0llOZ0dFtq0Z4+7X6gMTN9vMvpe784
// SIG // cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFxBmoQtB1V
// SIG // M1izoXBm8qGCAs8wggI4AgEBMIH8oYHUpIHRMIHOMQsw
// SIG // CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
// SIG // MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
// SIG // b2Z0IENvcnBvcmF0aW9uMSkwJwYDVQQLEyBNaWNyb3Nv
// SIG // ZnQgT3BlcmF0aW9ucyBQdWVydG8gUmljbzEmMCQGA1UE
// SIG // CxMdVGhhbGVzIFRTUyBFU046RjdBNi1FMjUxLTE1MEEx
// SIG // JTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
// SIG // cnZpY2WiIwoBATAHBgUrDgMCGgMVALPJcNtFs5sQyojd
// SIG // S4Ye5mVl7rSooIGDMIGApH4wfDELMAkGA1UEBhMCVVMx
// SIG // EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
// SIG // ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
// SIG // dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
// SIG // bXAgUENBIDIwMTAwDQYJKoZIhvcNAQEFBQACBQDnDePl
// SIG // MCIYDzIwMjIxMTAzMTA1MzI1WhgPMjAyMjExMDQxMDUz
// SIG // MjVaMHQwOgYKKwYBBAGEWQoEATEsMCowCgIFAOcN4+UC
// SIG // AQAwBwIBAAICIy0wBwIBAAICEX0wCgIFAOcPNWUCAQAw
// SIG // NgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAK
// SIG // MAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkqhkiG9w0B
// SIG // AQUFAAOBgQBtPJzT9CgITjEWI/emvodiiPqxmkAXTd6p
// SIG // 4ZFyv8BBEEi59kBvFSIO+buv6xr5IEGo5seVCgTxcNhB
// SIG // Y2D+FbD7hSqElGH3z44Ar8KU0xvfNki4akdTUoC6bhaO
// SIG // GxuyDn9FCuSd4pDdXmmAHjqyHQzZRRPaldDPQYAdmXhr
// SIG // pEQ/tTGCBA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVT
// SIG // MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
// SIG // ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
// SIG // YXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0
// SIG // YW1wIFBDQSAyMDEwAhMzAAABpQDeCMRAB3FOAAEAAAGl
// SIG // MA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMx
// SIG // DQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIONG
// SIG // JMW0S8vNohaAxXAEmlCHQlvxsn3PLn1SEXZW6ZCvMIH6
// SIG // BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQguAo4cX5m
// SIG // BLGgrdgFPNyoYfuiR5cpNwe9L3zBzJQS3FwwgZgwgYCk
// SIG // fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
// SIG // Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
// SIG // TWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1N
// SIG // aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAA
// SIG // AaUA3gjEQAdxTgABAAABpTAiBCCWvGuhOi5B1m4GKI5q
// SIG // uh5BkfU+rqDdS8vXmWDWC8hftDANBgkqhkiG9w0BAQsF
// SIG // AASCAgByNupReUVg7SjrQS7GVN05zTYZZGfQxss2QUzB
// SIG // czo99rnRAAZZ/Uuo0hd5yoqOk1lRt4WILLwljPuKOWqL
// SIG // snbYvcHrMA+6rKdWde8WMzC5omyXDGB+H3X/fy6/PgnQ
// SIG // TYcq5Y2MQhi5kVbtQrQ8cnB5kXH154Cb/b+MbrL+AIj4
// SIG // bgHoD/cixBewv9k+d3twLAAonBx0YrmtgOnmWn0ofDgi
// SIG // xFuDDCzl7RAgjGsZ0tjy71zrXJgKshhjWMpjnOfSJqjZ
// SIG // vVFtDPswuzymD3l4WuKtaTivAzImybwpRVPGycHhue1K
// SIG // /Kgt6HQd58g9kGhudYYHTRh5fjuM2Qv1o0pvsbwNeZ+r
// SIG // 1xGazmT7VqGQ0ULOuQcNTEE8RS1OWB9VJOg0vbhBo3pC
// SIG // rE0r94lRSgqOje+OIDEZ0fGe+A4HrwQSg7mSd7u3dUJN
// SIG // xpKrWVQmex/utCtNvXAbDmrcBMbL6QbXCh807HIaLNsL
// SIG // C75tOdTIGjuQWkrpmTFDCPQcHy4nJ2GG/gV/lNIar4GT
// SIG // hlaDgDZAUIcF+NV9fGiyZCYpfSJxkhR2bQHJXJ4VjSUk
// SIG // koeK/hGMKlqw1eVQf+GlHKe2Af8wTySOZa+q+S6hFJGM
// SIG // yXgrW4IzW5wYdO0/y2ctrKrIKfeDvxrTCgb0Sts9bdbz
// SIG // pMiSEIRUgUOCrowp0Zmy7lOSaZTGEA==
// SIG // End signature block
