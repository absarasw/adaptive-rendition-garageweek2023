  /*
 * ADOBE CONFIDENTIAL
 *
 * Copyright 2022 Adobe. All rights reserved.
 *
 * NOTICE: All information contained herein is, and remains
 * the property of Adobe Incorporated and its suppliers,
 * if any. The intellectual and technical concepts contained
 * herein are proprietary to Adobe Incorporated and its
 * suppliers and may be covered by U.S. and Foreign Patents,
 * patents in process, and are protected by trade secret or copyright law.
 * Dissemination of this information or reproduction of this material
 * is strictly forbidden unless prior written permission is obtained
 * from Adobe Incorporated.
 */

const Busboy = require('busboy');
const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const fetch = require('node-fetch');

// import axios from 'axios';
// import FormData from 'form-data';
// import fs from 'fs';

const saveImageFilePath = '/Users/abhinavsaraswat/Desktop/';
const downloadUrl = 'https://images.pexels.com/photos/60597/dahlia-red-blossom-bloom-60597.jpeg?cs=srgb&dl=pexels-pixabay-60597.jpg&fm=jpg&_gl=1*1v7pi2k*_ga*MTM1Mjc3OTgzOS4xNjkwMzAxOTY2*_ga_8JE65Q40S6*MTY5MDMwMTk2Ni4xLjEuMTY5MDMwMTk4NC4wLjAuMA..'; // Replace with the URL of the binary file you want to download
//const targetFilePath = '/Users/abhinavsaraswat/Documents/wallpaper.jpeg'; // Replace with the destination file path on your system

const boundary = `--------------------------${Date.now()}`;
const imageFile = '/Users/abhinavsaraswat/Documents/wallpaper.jpeg';

// const jsonContent = '{\"graph\":{\"uri\":\"urn:graph:MultiDiffusion_outpaint_v3\"},\"params\":[{\"name\":\"gi_MODE\",\"value\":\"ginp\",\"type\":\"string\"},{\"name\":\"gi_SEED\",\"value\":[{\"name\":\"0\",\"value\":78841,\"type\":\"scalar\"},{\"name\":\"1\",\"value\":11158,\"type\":\"scalar\"},{\"name\":\"2\",\"value\":81232,\"type\":\"scalar\"},{\"name\":\"3\",\"value\":26310,\"type\":\"scalar\"}],\"type\":\"array\"},{\"name\":\"gi_NUM_STEPS\",\"value\":70,\"type\":\"scalar\"},{\"name\":\"gi_GUIDANCE\",\"value\":6,\"type\":\"scalar\"},{\"name\":\"gi_ENABLE_PROMPT_FILTER\",\"value\":true,\"type\":\"boolean\"},{\"name\":\"gi_OUTPUT_WIDTH\",\"value\":1408,\"type\":\"scalar\"},{\"name\":\"gi_OUTPUT_HEIGHT\",\"value\":1024,\"type\":\"scalar\"},{\"name\":\"gi_AR_SHIFT\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_AR_DILATE\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_ADVANCED\",\"value\":\"{\\\"enable_mts\\\":true}\",\"type\":\"string\"}],\"inputs\":{\"gi_IMAGE\":{\"id\":\"a8226039-72c4-4bf3-bae1-faaca5561347\",\"toStore\":{\"lifeCycle\":\"session\"},\"type\":\"image\"}},\"outputs\":{\"gi_GEN_IMAGE\":{\"type\":\"array\",\"expectedMimeType\":\"image\\/jpeg\",\"expectedArrayLength\":1,\"id\":\"74c63a43-1a7b-4d0b-b950-75f3ba97adf8\"},\"gi_GEN_STATUS\":{\"type\":\"array\",\"id\":\"0ef7ad76-9aa6-4f5c-b7dc-43bb0f3ed1fd\"}}}';

const jsonContent = '{\"graph\":{\"uri\":\"urn:graph:MultiDiffusion_outpaint_v3\"},\"params\":[{\"name\":\"gi_MODE\",\"value\":\"ginp\",\"type\":\"string\"},{\"name\":\"gi_SEED\",\"value\":[{\"name\":\"0\",\"value\":78841,\"type\":\"scalar\"}],\"type\":\"array\"},{\"name\":\"gi_NUM_STEPS\",\"value\":70,\"type\":\"scalar\"},{\"name\":\"gi_GUIDANCE\",\"value\":6,\"type\":\"scalar\"},{\"name\":\"gi_ENABLE_PROMPT_FILTER\",\"value\":true,\"type\":\"boolean\"},{\"name\":\"gi_OUTPUT_WIDTH\",\"value\":1408,\"type\":\"scalar\"},{\"name\":\"gi_OUTPUT_HEIGHT\",\"value\":1024,\"type\":\"scalar\"},{\"name\":\"gi_AR_SHIFT\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_AR_DILATE\",\"value\":0,\"type\":\"scalar\"},{\"name\":\"gi_ADVANCED\",\"value\":\"{\\\"enable_mts\\\":true}\",\"type\":\"string\"}],\"inputs\":{\"gi_IMAGE\":{\"id\":\"a8226039-72c4-4bf3-bae1-faaca5561347\",\"toStore\":{\"lifeCycle\":\"session\"},\"type\":\"image\"}},\"outputs\":{\"gi_GEN_IMAGE\":{\"type\":\"array\",\"expectedMimeType\":\"image\\/jpeg\",\"expectedArrayLength\":1,\"id\":\"74c63a43-1a7b-4d0b-b950-75f3ba97adf8\"},\"gi_GEN_STATUS\":{\"type\":\"array\",\"id\":\"0ef7ad76-9aa6-4f5c-b7dc-43bb0f3ed1fd\"}}}';

const accessToken = 'eyJhbGciOiJSUzI1NiIsIng1dSI6Imltc19uYTEta2V5LWF0LTEuY2VyIiwia2lkIjoiaW1zX25hMS1rZXktYXQtMSIsIml0dCI6ImF0In0.eyJpZCI6IjE2OTA1MjM5NTE3NDhfZjQ1MjJkMjUtMzFkYi00NTY5LWJkNjktM2VhYjFkNDI3NTY0X3V3MiIsInR5cGUiOiJhY2Nlc3NfdG9rZW4iLCJjbGllbnRfaWQiOiJjbGlvLXBsYXlncm91bmQtd2ViIiwidXNlcl9pZCI6IjE3MUExREVBNjQ5MkIyRDEwQTQ5NUZCQ0BkMTEzMWVjNjY0NzczNjg1NDk1YzY4LmUiLCJzdGF0ZSI6IntcImpzbGlidmVyXCI6XCJ2Mi12MC4zMS4wLTItZzFlOGE4YThcIixcIm5vbmNlXCI6XCI1MjQxNzUyNjM4NTQxMTAxXCJ9IiwiYXMiOiJpbXMtbmExIiwiYWFfaWQiOiIzNDhGNzIzMTVBMTU0MERCMEE0OTVFRkVAYWRvYmUuY29tIiwiY3RwIjowLCJmZyI6IlhVVlFQWkpHWFBQNzRQNktHT1FWM1hBQTRBPT09PT09Iiwic2lkIjoiMTY5MDUyMzk1MTEzMV83YjhmMzE3MS0xMjQ5LTQyNDktOGMxNC02ZDQ4NjkxZjIyNGFfdXcyIiwibW9pIjoiMjNhMGQxNGUiLCJwYmEiOiJNZWRTZWNOb0VWLExvd1NlYyIsImV4cGlyZXNfaW4iOiI4NjQwMDAwMCIsInNjb3BlIjoiQWRvYmVJRCxvcGVuaWQsZmlyZWZseV9hcGkiLCJjcmVhdGVkX2F0IjoiMTY5MDUyMzk1MTc0OCJ9.UMk75HRcvuzovLIKauYH_T92QMExkTZJ6BUdoMLld9PK2wU_8WWf5I8ziwDIi8eujW2_p-P-nyqkcGq5Qwj2as70Ao0Hbt5fVVrpM8nhVwsc34fSCbGWkaSZXsR_RxOUlbYfDRwP__XxyyVZYtj5OEt-El4ZN7FqVKtFxCcUJQNsjppfQvzy5OiVPUJkJ53UstjRTZddtrjf0_pJmoFioIkGMYsOLK9hQ1hdEeugukp_FWPc3bhpZRf9ZkmlTB4WGZ-imC1-32tpdUgX_hwn486wi7hhgCcqUjrSmB9T6yT_dDbTLvRRBrftQpItTkD4O95H2QyRcM6rPRTzziofrQ';

const apiEndpoint = 'https://firefly.adobe.io/spl';

async function downloadAsset(downloadUrl, targetFilePath) {
  axios({
    method: 'get',
    url: downloadUrl,
    responseType: 'arraybuffer', // Important: Set the responseType to 'arraybuffer' to handle binary data
  })
    .then((response) => {
      fs.writeFileSync(targetFilePath, response.data);
      console.log('File downloaded successfully!');
    })
    .catch((error) => {
      console.error('Error downloading the file:', error.message);
    });
}



async function sendMultipartRequest(imageFile) {
  const formData = new FormData();
  formData.boundary = boundary;
  formData.append('request', jsonContent, {
    contentType: 'application/json',
  });

  const imgUrl = 'https://main--adaptive-rendition-garageweek2023--absarasw.hlx.live/wallpaper.jpeg';

  const response = await fetch(imgUrl);

  if (!response.ok) {
    throw new Error('Failed to download the image.');
  }

  const imageBlob = await response.blob();

  formData.append('gi_IMAGE', imageBlob.stream(), {
    filename: 'image.jpg',
    contentType: 'image/jpeg',
  });

  /*formData.append('gi_IMAGE', fs.createReadStream(imageFile), {
    filename: 'image.jpg',
    contentType: 'image/jpeg',
  });*/




  try {
    const response = await axios.post(apiEndpoint, formData, {
      headers: {
        ...formData.getHeaders(),
        'x-api-key': 'clio-playground-web',
        Authorization: `Bearer ${accessToken}`,
        'x-session-id': 'b2382afb-1324-44be-844e-63ef60e77cbf',
        'Accept-Encoding': 'gzip, deflate, br',
      },
    });
    console.log(response.headers['content-type']);
    const contentType = response.headers['content-type'];

    if (contentType && contentType.includes('multipart/form-data')) {
      // fs.writeFileSync(saveImageFilePath, Buffer.from(response.data));

      return processMultipartEntity_1(response, saveImageFilePath);
    } else {
      // Handle other types of responses
      console.log(response.data);
    }
  } catch (error) {
    console.error(error);
  }
}

async function processMultipartEntity_1(entity, saveImageFilePath) {
  const contentType = entity.headers.get('content-type');
  const boundary = contentType.split('boundary=')[1];
  const multipart = require('parse-multipart-data');
  const getStream = (await import('get-stream')).default;

  const collectChunks = async (readable) => Buffer.from(await getStream(readable));

  const bodyContent = await collectChunks(entity.data);

  const parts = multipart.parse(bodyContent, boundary);

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if(part.name && part.name.includes('gi_GEN_IMAGE')) {
      const data = part.data;

      const outputFileName = saveImageFilePath + 'image_1.jpeg';
      const buffer = Buffer.from(data);

      return data;
      //fs.createWriteStream(outputFileName).write(buffer);

      /*fs.writeFile(saveImageFilePath + 'image_1.jpeg', data, {encoding: 'utf-8'}, function (err) {
        if (err) throw err;
        console.log('It\'s saved!');
      });*/

    }
    // will be: { filename: 'A.txt', type: 'text/plain', data: <Buffer 41 41 41 41 42 42 42 42> }
  }
}

async function processMultipartEntity(entity, saveImageFilePath) {
  const contentType = entity.headers.get('content-type');
  const boundary = contentType.split('boundary=')[1];

  const formData = new FormData();

  const partHeaders = [];
  let partData = Buffer.from([]);

  // Change the import statement to dynamic import
  const getStream = (await import('get-stream')).default;

  const collectChunks = async (readable) => Buffer.from(await getStream(readable));

  const bodyContent = await collectChunks(entity.data);

  let offset = 0;
  while (offset < bodyContent.length) {
    const boundaryIndex = bodyContent.indexOf(boundary, offset);
    if (boundaryIndex !== -1) {
      if (partHeaders.length > 0) {
        formData.append('part', partData, {
          header: partHeaders.join('\r\n'),
        });
      }

      partHeaders.length = 0;
      partData = Buffer.from([]);
      offset = boundaryIndex + boundary.length;
    } else if (partHeaders.length === 0) {
      const lineEndIndex = bodyContent.indexOf('\r\n', offset);
      if (lineEndIndex !== -1) {
        partHeaders.push(bodyContent.slice(offset, lineEndIndex).toString());
        offset = lineEndIndex + 2;
      } else {
        break; // Not enough data to parse headers, wait for more chunks
      }
    } else {
      const nextBoundaryIndex = bodyContent.indexOf(boundary, offset);
      if (nextBoundaryIndex !== -1) {
        partData = Buffer.concat([partData, bodyContent.slice(offset, nextBoundaryIndex)]);
        offset = nextBoundaryIndex;
      } else {
        partData = Buffer.concat([partData, bodyContent.slice(offset)]);
        break;
      }
    }
  }

  await new Promise((resolve, reject) => {
    const busboy = Busboy({ headers: { 'content-type': contentType } });

    busboy.on('file', (fieldname, file, filename) => {
      const filePath = `${saveImageFilePath}firefly_${fieldname}.jpeg`;
      console.log(`Saving file: ${filePath}`);

      const writeStream = fs.createWriteStream(filePath);
      file.pipe(writeStream);
      file.on('end', () => resolve());
    });

    busboy.on('finish', resolve);
    busboy.end(formData.getBuffer());
  });
}

async function saveFile(file, filePath) {
  const fileContent = await fs.promises.readFile(file.path);
  await fs.promises.writeFile(filePath, fileContent);
}

  module.exports = sendMultipartRequest;
