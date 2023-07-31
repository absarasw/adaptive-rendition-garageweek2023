import { PublicClientApplication } from './msal-browser-2.14.2.js';
import { Document, Paragraph, Packer, HeadingLevel } from 'docx';
import { saveAs } from 'file-saver';
import fs from "fs";

const graphURL = 'https://graph.microsoft.com/v1.0';
const baseURI = 'https://graph.microsoft.com/v1.0/drives/b!9IXcorzxfUm_iSmlbQUd2rvx8XA-4zBAvR2Geq4Y2sZTr_1zgLOtRKRA81cvIhG1/root:/fcbayern';
const driveIDGlobal = 'b!9IXcorzxfUm_iSmlbQUd2rvx8XA-4zBAvR2Geq4Y2sZTr_1zgLOtRKRA81cvIhG1';
const folderID = '01DF7GY26OP3TGBGEQJVHLMHAETWTTMQEG';
let connectAttempts = 0;
let accessToken;

const orgName = 'absarasw';
const repoName = 'adaptive-rendition-garageweek2023';
const ref = 'main';
const path = 'de/spiele/profis/bundesliga/2022-2023/sv-werder-bremen-fc-bayern-muenchen-06-05-2023/liveticker';
const mockNotificationService = 'https://288650-257ambermackerel.adobeio-static.net/api/v1/web/brandads/getads';
const test = `abc`;

const sp = {
    clientApp: {
        auth: {
            clientId: '0b2504a7-ae7a-4cd3-adab-213f7bfb0c25',
            authority: 'https://login.microsoftonline.com/fa7b1b5a-7b34-4387-94ae-d2c178decee1',
        },
    },
    login: {
        redirectUri: '/tools/sidekick/spauth.html',
    },
};



//const downloadUrl = 'https://images.pexels.com/photos/60597/dahlia-red-blossom-bloom-60597.jpeg?cs=srgb&dl=pexels-pixabay-60597.jpg&fm=jpg&_gl=1*1v7pi2k*_ga*MTM1Mjc3OTgzOS4xNjkwMzAxOTY2*_ga_8JE65Q40S6*MTY5MDMwMTk2Ni4xLjEuMTY5MDMwMTk4NC4wLjAuMA..'; // Replace with the URL of the binary file you want to download
const downloadUrl = 'https://main--adaptive-rendition-garageweek2023--absarasw.hlx.live/wallpaper.jpeg';
const hostname = 'https://adaptive-renditions.aem-screens.com';
const folderPath = 'content/screens/assets';


// const jsonContent = '{\"graph\":{\"uri\":\"urn:graph:MultiDiffusion_outpaint_v3\"},\"params\":[{\"name\":\"gi_MODE\",\"value\":\"ginp\",\"type\":\"string\"},{\"name\":\"gi_SEED\",\"value\":[{\"name\":\"0\",\"value\":78841,\"type\":\"scalar\"},{\"name\":\"1\",\"value\":11158,\"type\":\"scalar\"},{\"name\":\"2\",\"value\":81232,\"type\":\"scalar\"},{\"name\":\"3\",\"value\":26310,\"type\":\"scalar\"}],\"type\":\"array\"},{\"name\":\"gi_NUM_STEPS\",\"value\":70,\"type\":\"scalar\"},{\"name\":\"gi_GUIDANCE\",\"value\":6,\"type\":\"scalar\"},{\"name\":\"gi_ENABLE_PROMPT_FILTER\",\"value\":true,\"type\":\"boolean\"},{\"name\":\"gi_OUTPUT_WIDTH\",\"value\":1408,\"type\":\"scalar\"},{\"name\":\"gi_OUTPUT_HEIGHT\",\"value\":1024,\"type\":\"scalar\"},{\"name\":\"gi_AR_SHIFT\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_AR_DILATE\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_ADVANCED\",\"value\":\"{\\\"enable_mts\\\":true}\",\"type\":\"string\"}],\"inputs\":{\"gi_IMAGE\":{\"id\":\"a8226039-72c4-4bf3-bae1-faaca5561347\",\"toStore\":{\"lifeCycle\":\"session\"},\"type\":\"image\"}},\"outputs\":{\"gi_GEN_IMAGE\":{\"type\":\"array\",\"expectedMimeType\":\"image\\/jpeg\",\"expectedArrayLength\":1,\"id\":\"74c63a43-1a7b-4d0b-b950-75f3ba97adf8\"},\"gi_GEN_STATUS\":{\"type\":\"array\",\"id\":\"0ef7ad76-9aa6-4f5c-b7dc-43bb0f3ed1fd\"}}}';

const jsonContent = '{\"graph\":{\"uri\":\"urn:graph:MultiDiffusion_outpaint_v3\"},\"params\":[{\"name\":\"gi_MODE\",\"value\":\"ginp\",\"type\":\"string\"},{\"name\":\"gi_SEED\",\"value\":[{\"name\":\"0\",\"value\":78841,\"type\":\"scalar\"}],\"type\":\"array\"},{\"name\":\"gi_NUM_STEPS\",\"value\":70,\"type\":\"scalar\"},{\"name\":\"gi_GUIDANCE\",\"value\":6,\"type\":\"scalar\"},{\"name\":\"gi_ENABLE_PROMPT_FILTER\",\"value\":true,\"type\":\"boolean\"},{\"name\":\"gi_OUTPUT_WIDTH\",\"value\":1408,\"type\":\"scalar\"},{\"name\":\"gi_OUTPUT_HEIGHT\",\"value\":1024,\"type\":\"scalar\"},{\"name\":\"gi_AR_SHIFT\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_AR_DILATE\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_ADVANCED\",\"value\":\"{\\\"enable_mts\\\":true}\",\"type\":\"string\"}],\"inputs\":{\"gi_IMAGE\":{\"id\":\"a8226039-72c4-4bf3-bae1-faaca5561347\",\"toStore\":{\"lifeCycle\":\"session\"},\"type\":\"image\"}},\"outputs\":{\"gi_GEN_IMAGE\":{\"type\":\"array\",\"expectedMimeType\":\"image\\/jpeg\",\"expectedArrayLength\":1,\"id\":\"74c63a43-1a7b-4d0b-b950-75f3ba97adf8\"},\"gi_GEN_STATUS\":{\"type\":\"array\",\"id\":\"0ef7ad76-9aa6-4f5c-b7dc-43bb0f3ed1fd\"}}}';

const jsonContentPortrait = '{\"graph\":{\"uri\":\"urn:graph:MultiDiffusion_outpaint_v3\"},\"params\":[{\"name\":\"gi_MODE\",\"value\":\"ginp\",\"type\":\"string\"},{\"name\":\"gi_SEED\",\"value\":[{\"name\":\"0\",\"value\":78841,\"type\":\"scalar\"}],\"type\":\"array\"},{\"name\":\"gi_NUM_STEPS\",\"value\":70,\"type\":\"scalar\"},{\"name\":\"gi_GUIDANCE\",\"value\":6,\"type\":\"scalar\"},{\"name\":\"gi_ENABLE_PROMPT_FILTER\",\"value\":true,\"type\":\"boolean\"},{\"name\":\"gi_OUTPUT_WIDTH\",\"value\":1024,\"type\":\"scalar\"},{\"name\":\"gi_OUTPUT_HEIGHT\",\"value\":1408,\"type\":\"scalar\"},{\"name\":\"gi_AR_SHIFT\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_AR_DILATE\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_ADVANCED\",\"value\":\"{\\\"enable_mts\\\":true}\",\"type\":\"string\"}],\"inputs\":{\"gi_IMAGE\":{\"id\":\"a8226039-72c4-4bf3-bae1-faaca5561347\",\"toStore\":{\"lifeCycle\":\"session\"},\"type\":\"image\"}},\"outputs\":{\"gi_GEN_IMAGE\":{\"type\":\"array\",\"expectedMimeType\":\"image\\/jpeg\",\"expectedArrayLength\":1,\"id\":\"74c63a43-1a7b-4d0b-b950-75f3ba97adf8\"},\"gi_GEN_STATUS\":{\"type\":\"array\",\"id\":\"0ef7ad76-9aa6-4f5c-b7dc-43bb0f3ed1fd\"}}}';

const firefly_accessToken = 'eyJhbGciOiJSUzI1NiIsIng1dSI6Imltc19uYTEta2V5LWF0LTEuY2VyIiwia2lkIjoiaW1zX25hMS1rZXktYXQtMSIsIml0dCI6ImF0In0.eyJpZCI6IjE2OTA3MjQ4NTUwODZfMTI0ZDM2NGQtNDA4Ny00NjIzLWI3YjUtMTk2MjhkMmUyZWZmX3V3MiIsInR5cGUiOiJhY2Nlc3NfdG9rZW4iLCJjbGllbnRfaWQiOiJjbGlvLXBsYXlncm91bmQtd2ViIiwidXNlcl9pZCI6IjgyRkMxRUNENjMxQzMzRDUwQTQ5NUMzRUA3ZWViMjBmODYzMWMwY2I3NDk1YzA2LmUiLCJzdGF0ZSI6IntcInNlc3Npb25cIjpcImh0dHBzOi8vaW1zLW5hMS5hZG9iZWxvZ2luLmNvbS9pbXMvc2Vzc2lvbi92MS9NekUzTURFMlpURXRabVF4TkMwMFlUUTVMV0UwWTJFdFpqRTBaV0UwTWpRMk1ERmpMUzA0TWtaRE1VVkRSRFl6TVVNek0wUTFNRUUwT1RWRE0wVkFOMlZsWWpJd1pqZzJNekZqTUdOaU56UTVOV013Tmk1bFwifSIsImFzIjoiaW1zLW5hMSIsImFhX2lkIjoiMzQ4RjcyMzE1QTE1NDBEQjBBNDk1RUZFQGFkb2JlLmNvbSIsImN0cCI6MCwiZmciOiJYVTRCVzNVT1hQUDdNUDZLR09RVjNYQUFQQT09PT09PSIsInNpZCI6IjE2OTA1NDY5NTYwOTVfNjAzZDM2NDItMzc0Ny00Y2RkLTk5YTctZWUxYjg4NzcxMGYxX3V3MiIsIm1vaSI6IjIzNzJkNTJhIiwicGJhIjoiTWVkU2VjTm9FVixMb3dTZWMiLCJleHBpcmVzX2luIjoiODY0MDAwMDAiLCJzY29wZSI6IkFkb2JlSUQsb3BlbmlkLGZpcmVmbHlfYXBpIiwiY3JlYXRlZF9hdCI6IjE2OTA3MjQ4NTUwODYifQ.HrAsRwkyGp3dDpDkOtYsgfh1fm4KHj1UbXL_9Ta4dC2T7jVt2OnbeJa2ZIhN5dL7PzorSdhlAHVlrwdfQz-S5hIfPwvtYERd-biUp95dZpuES3jtFEWghugfKcMTOHIYTYvweRCOuAOroAl_mmavv6RBTT_VL13DoGoL4_i7xD0zSD2k2utJ8m8gtLb3hCGttHbDhq4PNhsjaOnLmuXI6cOcl-HJ0MTXHPkGCJ62unf8Nw-CINcw3pWFMWuBuuTnd5sB835AJOXhRzfZDeVvxOTCjaEx9Wgz6Eql8A_rPJfJCZ5Iy7Gr2KdbiJA6EkeCcDY6YqqoP-pFuuY2iAO7Jw';
const apiEndpoint = 'https://firefly.adobe.io/spl';


async function getAdaptiveRendition(imagePath, jsonContent) {
    const formData = new FormData();

    const downloadUrl = hostname + "/" + folderPath + "/" + imagePath + ".jpg";
    const imageBuffer = await fetch(downloadUrl).then((response) => response.arrayBuffer());
    const imageBlob = new Blob([imageBuffer], { type: 'image/jpeg' });
    formData.append('gi_IMAGE', imageBlob, 'blob');
    const jsonBlob = new Blob([jsonContent], { type: 'application/json' });
    formData.append('request', jsonBlob);
    try {
        const response = await fetch(apiEndpoint, {
            method: 'POST',
            headers: {
                Authorization: `Bearer ${firefly_accessToken}`,
                'x-api-key': 'clio-playground-web',
                'x-session-id': 'b2382afb-1324-44be-844e-63ef60e77cbf',
                'Accept-Encoding': 'gzip, deflate, br',
                'access-control-allow-origin': '*',
                'x-request-id': '477f762a-02c9-4f2f-ae65-ff8fcc823111',
            },
            body: formData,
        });
        const responseData = await response.formData();
        return responseData.get('gi_GEN_IMAGE_0');
    } catch (error) {
        console.error(error);
    }
}

async function sendMultipartRequest(imagePath) {
    /*const formData = new FormData();

    const downloadUrl = hostname + "/" + folderPath + "/" + imagePath + ".jpg";
    const imageBuffer = await fetch(downloadUrl).then((response) => response.arrayBuffer());
    const imageBlob = new Blob([imageBuffer], { type: 'image/jpeg' });
    formData.append('gi_IMAGE', imageBlob, 'blob');
    const jsonBlob = new Blob([jsonContent], { type: 'application/json' });
    formData.append('request', jsonBlob);
    try {
        const response = await fetch(apiEndpoint, {
            method: 'POST',
            headers: {
                Authorization: `Bearer ${firefly_accessToken}`,
                'x-api-key': 'clio-playground-web',
                'x-session-id': 'b2382afb-1324-44be-844e-63ef60e77cbf',
                'Accept-Encoding': 'gzip, deflate, br',
                'access-control-allow-origin': '*',
                'x-request-id': '477f762a-02c9-4f2f-ae65-ff8fcc823111',
            },
            body: formData,
        });
        const responseData = await response.formData();
        const blob = responseData.get('gi_GEN_IMAGE_0');*/
        const landscapeBlob = await getAdaptiveRendition(imagePath, jsonContent);
        const portraitBlob = await getAdaptiveRendition(imagePath, jsonContentPortrait);

        const renditionFolderId = await createFolder(imagePath + "_renditions");

        const landscapeRenditionName = imagePath + "_landscape";
        const portraitRenditionName = imagePath + "_portrait";
        await uploadImageFromBlob(landscapeBlob, renditionFolderId, landscapeRenditionName);
        await uploadImageFromBlob(portraitBlob, renditionFolderId, portraitRenditionName);
        const r1 = folderPath + "/" + imagePath + "_renditions/" + imagePath + "_landscape.jpeg";
        const r2 = folderPath + "/" + imagePath + "_renditions/" + imagePath + "_portrait.jpeg";
        await quickpublish(r1);
        await quickpublish(r2);

}

async function uploadImageFromBlob(imageBlob, folderID, imageName) {
    const { size, type } = imageBlob;
    console.log(`IMG1 Type: ${type}\n IMG Size: ${size}`);

    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveIDGlobal}/items/${folderID}:/${imageName}.jpeg:/content`;

    const uploadResponse = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': imageBlob.type
        },
        body: imageBlob
    });
    if (uploadResponse.ok) {
        const response = await uploadResponse.json();
        console.log('Image has been uploaded');
    } else {
        console.log('here 4');
    }
}





export async function connect(callback) {
    const publicClientApplication = new PublicClientApplication(sp.clientApp);

    const accounts = publicClientApplication.getAllAccounts();

    if (accounts.length === 0) {
        // User is not logged in, show the login popup
        await publicClientApplication.loginPopup(sp.login);
    }

    const account = publicClientApplication.getAllAccounts()[0];

    const accessTokenRequest = {
        scopes: ['files.readwrite', 'sites.readwrite.all'],
        account,
    };

    try {
        const res = await publicClientApplication.acquireTokenSilent(accessTokenRequest);
        accessToken = res.accessToken;
        if (callback) await callback();
    } catch (error) {
        // Acquire token silent failure, and send an interactive request
        if (error.name === 'InteractionRequiredAuthError') {
            try {
                const res = await publicClientApplication.acquireTokenPopup(accessTokenRequest);
                // Acquire token interactive success
                accessToken = res.accessToken;
                if (callback) await callback();
            } catch (err) {
                connectAttempts += 1;
                if (connectAttempts === 1) {
                    // Retry to connect once
                    connect(callback);
                }
                // Give up
                throw new Error(`Cannot connect to Sharepoint: ${err.message}`);
            }
        }
    }
}

function validateConnnection() {
    if (!accessToken) {
        throw new Error('You need to sign-in first');
    }
}

function getRequestOption() {
    validateConnnection();

    const bearer = `Bearer ${accessToken}`;
    const headers = new Headers();
    headers.append('Authorization', bearer);

    return {
        method: 'GET',
        headers,
    };
}

async function isAssetAvailable(i) {
    const resourcePath = hostname + "/" + folderPath + "/" + i + ".jpg";
    const resp = await fetch(
        resourcePath,
        { method: 'HEAD' }
    );

    return resp.ok;
}

export async function PublishAndNotify() {
    // const quickPublish = await quickpublish();
    // if (quickPublish === 'published') {
    //     return 'updated';
    // }
    //await uploadDocumentFile(folderID);
    //await sendMultipartRequest();
    //const parentFolderPath = 'abhinavscreens/' + folderPath;
    let i = 1;
    while(await isAssetAvailable(i)) {
        await sendMultipartRequest(i);
        i++;
    }
}


async function uploadDocumentFile(folderId) {

    const doc = new Document({
        sections: [
            {
                properties: {},
                children: [
                    new Paragraph({
                        text: "This paragraph will be in my new document",
                        heading: HeadingLevel.HEADING_1, // Set appropriate heading level
                    }),
                ],
            },
        ],
    });

    try {
        const buffer = await Packer.toBuffer(doc);
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        //saveAs(blob, `first.docx`);
        const fileName = 'first.docx';
        const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveIDGlobal}/items/${folderId}:/${fileName}:/content`;

        const uploadResponse = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            },
            body: blob
        });
        if (uploadResponse.ok) {
            const response = await uploadResponse.json();
            console.log('Document has been uploaded1');
        } else {
            console.log('here 4');
        }


    } catch (error) {
        console.error("Error creating or saving the document:", error);
    }
}

async function uploadImage(folderId) {
    const imageUrl = 'https://raw.githubusercontent.com/anagarwa/adobe-screens-brandads/main/content/dam/ads/mdsrimages/ad4/1.png';
    // Download the image from the URL
    const response = await fetch(imageUrl);
    if (!response.ok) {
        throw new Error('Failed to download the image.');
    }

    const imageBlob = await response.blob();
    const { size, type } = imageBlob;
    console.log(`IMG1 Type: ${type}\n🌌 IMG Size: ${size}`);

    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveIDGlobal}/items/${folderId}:/${getImageFileName(imageUrl)}:/content`;

    const uploadResponse = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': imageBlob.type
        },
        body: imageBlob
    });
    if (uploadResponse.ok) {
        const response = await uploadResponse.json();
        console.log('Image has been uploaded');
    } else {
        console.log('here 4');
    }
}

function getImageFileName(imageUrl) {
    const parts = imageUrl.split('/');
    return parts[parts.length - 1];
}

async function createFolder(folderName) {
    validateConnnection();
    const folderData = {
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename"
    };
    const createUrl = `https://graph.microsoft.com/v1.0/drives/${driveIDGlobal}/items/${folderID}/children`;
    const createResponse = await fetch(createUrl, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(folderData)
    });
    const response = await createResponse.json();
    if (createResponse.ok) {
        console.log("folder is created" + response.id);
    } else {
        throw new Error('Failed to create folder');
    }
    return response.id;
}

async function getFolderID(parentFolderPath) {
    try {
        validateConnnection();
        const options = getRequestOption();
        //const parentFolderPath = 'abhinavscreens/content/screens/assets';
        const getByPathUrl = `https://graph.microsoft.com/v1.0/drives/${driveIDGlobal}/root:/${parentFolderPath}:/`;
        const driveResponse = await fetch(getByPathUrl, options);
        const response = await driveResponse.json();
        const folderId = response.id;
        console.log("folder id is " + folderId);
        return folderId;
    } catch (error) {
        throw new Error('Failed to retrieve folder ID');
    }
}

async function quickpublish(path) {
    console.log('in quick publish ' + path);
    console.log(`Quick Publish Started ${new Date().toLocaleString()}`);

    let response;
    const options = {
        method: 'POST',
    };

    response = await fetch(`https://admin.hlx.page/preview/${orgName}/${repoName}/${ref}/${path}`, options);

    if (response.ok) {
        console.log(`Document Previewed at ${new Date().toLocaleString()}`);
    } else {
        throw new Error(`Could not previewed. Status: ${response.status}`);
    }

    response = await fetch(`https://admin.hlx.page/live/${orgName}/${repoName}/${ref}/${path}`, options);

    if (response.ok) {
        console.log(`Document Published at ${new Date().toLocaleString()}`);
    } else {
        throw new Error(`Could not published. Status: ${response.status}`);
    }

    response = await fetch(`https://admin.hlx.page/cache/${orgName}/${repoName}/${ref}/${path}`, options);

    if (response.ok) {
        console.log(`Purge cache ${new Date().toLocaleString()}`);
    } else {
        throw new Error(`Could not purge cache. Status: ${response.status}`);
    }
}

async function quickpublish_1(path) {
    console.log('in quick publish');
    console.log(`Quick Publish Started ${new Date().toLocaleString()}`);

    let response;
    const options = {
        method: 'POST',
    };

    response = await fetch(`https://admin.hlx.page/preview/${orgName}/${repoName}/${ref}/${path}`, options);

    if (response.ok) {
        console.log(`Document Previewed at ${new Date().toLocaleString()}`);
    } else {
        throw new Error(`Could not previewed. Status: ${response.status}`);
    }

    response = await fetch(`https://admin.hlx.page/live/${orgName}/${repoName}/${ref}/${path}`, options);

    if (response.ok) {
        console.log(`Document Published at ${new Date().toLocaleString()}`);
    } else {
        throw new Error(`Could not published. Status: ${response.status}`);
    }

    response = await fetch(`https://admin.hlx.page/cache/${orgName}/${repoName}/${ref}/${path}`, options);

    if (response.ok) {
        console.log(`Purge cache ${new Date().toLocaleString()}`);
    } else {
        throw new Error(`Could not purge cache. Status: ${response.status}`);
    }

    let fileId = localStorage.getItem('fileId');
    if (!fileId) {
        fileId = await getFileId();
        localStorage.setItem('fileId', fileId);
    }

    const driveId = driveIDGlobal;

    const sheetName = 'notifications';
    const lastRow = 0;
    let entryRowExcel = -1;
    const excelData = await getExcelData(driveId, fileId, sheetName);

    const livetickerurl = `https://${ref}--${repoName}--${orgName}.hlx.page/${path}`;

    const liveTickerResponse = await fetch(livetickerurl);
    const liveTickerHtml = await liveTickerResponse.text();
    console.log(liveTickerHtml);
    const parser = new DOMParser();
    const doc = parser.parseFromString(liveTickerHtml, 'text/html');

    const jsonArray = [];
    const eventElements = doc.querySelectorAll('.goal, .whistle');
    for (let j = 0; j < eventElements.length; j++) {
        const eventElement = eventElements[j];
        const jsonObject = {};
        jsonObject.eventType = eventElement.classList;
        const divElements = eventElement.querySelectorAll(':scope > div');
        for (let i = 0; i < divElements.length; i++) {
            const keyValueDiv = divElements[i].querySelectorAll('div');
            const key = keyValueDiv[0].textContent.trim().toLowerCase().replace(' ', '_');
            const value = keyValueDiv[1].textContent.trim();
            jsonObject[key] = value;
        }
        if (jsonObject.push === 'yes' || jsonObject.push === 'true') {
            // todo code to confirm if it has been updated in excel if not send notification and update excel

            for (let row = 0; row < excelData.values.length; row++) {
                if (excelData.values[row][0].toString().trim() === jsonObject.id.toString().trim()) {
                    // event already exists in Excel
                    break;
                }
                if (!excelData.values[row][0]) {
                    entryRowExcel = row + 2;

                    // sending notification data to notification service
                    const notificationResponse = await fetch(mockNotificationService, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                        },
                        body: JSON.stringify(jsonObject),
                    });

                    if (notificationResponse.ok) {
                        console.log(`Notification of ${jsonObject.id}  sent at ${new Date().toLocaleString()}`);
                        jsonArray.push(
                            {
                                id: jsonObject.id,
                                notificationData: JSON.stringify(jsonObject),
                            },
                        );
                    }
                    break;
                }
            }
        }
    }

    if (jsonArray.length > 0) {
        const addEntriesResponse = await addEntriesToExcel(driveId, fileId, sheetName, entryRowExcel, jsonArray);
    }
    return 'published';
}

async function getExcelData(driveId, fileId, sheetName) {
    const endpoint = `/drives/${driveId}/items/${fileId}/workbook/worksheets('${sheetName}')/range(address='A2:A100')`;

    validateConnnection();

    const options = getRequestOption();
    options.method = 'GET';
    options.headers.append('Content-Type', 'application/json');

    const response = await fetch(`${graphURL}${endpoint}`, options);

    if (response.ok) {
        const searchResults = await response.json();
        return searchResults;
    }

    throw new Error(`Could not add entries to Excel file. Status: ${response.status}`);
}

async function getDriveId() {
    try {
        validateConnnection();
        const options = getRequestOption();

        const driveResponse = await fetch('https://graph.microsoft.com/v1.0/me/drive', options);
        const driveData = await driveResponse.json();
        const driveId = driveData.id;

        return driveId;
    } catch (error) {
        throw new Error('Failed to retrieve drive ID');
    }
}

async function getFileId() {
    const endpoint = `${baseURI}/matchdata/pushnotifications.xlsx`;

    validateConnnection();

    const options = getRequestOption();
    options.headers.append('Content-Type', 'application/json');
    options.method = 'GET';

    const response = await fetch(`${endpoint}`, options);

    if (response.ok) {
        const file = await response.json();
        return file.id;
    }

    throw new Error(`Could not retrieve file ID. Status: ${response.status}`);
}

async function addEntriesToExcel(driveId, fileId, sheetName, entryRow, entries) {
    const lastRow = entryRow + (entries.length - 1);
    const endpoint = `/drives/${driveId}/items/${fileId}/workbook/worksheets('${sheetName}')/range(address='A${entryRow}:B${lastRow}')`;

    const requestBody = {
        values: entries.map((entry) => [entry.id, entry.notificationData]),
    };

    validateConnnection();

    const options = getRequestOption();
    options.method = 'PATCH';
    options.headers.append('Content-Type', 'application/json');
    options.body = JSON.stringify(requestBody);

    const response = await fetch(`${graphURL}${endpoint}`, options);

    if (response.ok) {
        return response.json();
    }

    throw new Error(`Could not add entries to Excel file. Status: ${response.status}`);
}
