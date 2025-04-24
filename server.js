const express = require('express');
const cors = require('cors');
require('dotenv').config()

const app = express();
app.use(express.json());
app.use(cors({
    origin: 'http://localhost:5173',
    credentials: true,
}));


// refresh graphApitokens
async function refreshTokens(refreshToken) {
    const { CLIENT_ID, CLIENT_SECRET} = process.env
    const url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

    const params = new URLSearchParams();
    params.append('client_id', CLIENT_ID);
    params.append('grant_type', 'refresh_token');
    params.append('client_secret', CLIENT_SECRET);
    params.append('refresh_token', refreshToken);

    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params.toString()
    });
    const data = await response.json();
    return {
        graphApiToken: data.access_token, 
        refreshToken: data.refresh_token
    }
}

// Get graphApi tokens
async function getTokenOnBehalfOf(userAccessToken) {
    const { CLIENT_ID, CLIENT_SECRET} = process.env
    const url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

    const params = new URLSearchParams();
    params.append('assertion', userAccessToken);
    params.append('client_id', CLIENT_ID);
    params.append('scope', 'Files.Read.All Sites.Read.All Sites.ReadWrite.All');
    params.append('grant_type', 'urn:ietf:params:oauth:grant-type:jwt-bearer');
    params.append('client_secret', CLIENT_SECRET);
    params.append('requested_token_use', 'on_behalf_of');
    
    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params.toString()
    });
    
    const data = await response.json();
    return {
        graphApiToken: data.access_token, 
        refreshToken: data.refresh_token
    }
}

// SHAREPOINT FETCH LOGIC

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function fetchSharePoint(url, accessToken) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    });
    if (!res.ok) {
      throw new Error(`Failed to fetch ${url}: ${res.statusText}`);
    }
    return res.json();
  }

async function traverseFolder(driveId, folderId, childrenArray, accessToken) {
    let url = `https://graph.microsoft.com/v1.0/sites/chapterworks.sharepoint.com/drives/${driveId}/items/${folderId}/children`;
    const { value: items } = await fetchSharePoint(url, accessToken)
    
    if (!items) return;
    
    for (const item of items) {
        if (item.folder) {
            let folderNode = {
                id: item.id,
                name: item.name,
                children: []
            };

            childrenArray.push(folderNode);

            if (Number(item.folder.childCount) > 0) {
                await sleep(120);
                await traverseFolder(driveId, folderNode.id, folderNode.children, accessToken);
            }
        }
    }
}

async function fetchFolderTree(accessToken) {
    const folderTree = []
    const {value: drives} = await fetchSharePoint(
        `https://graph.microsoft.com/v1.0/sites/chapterworks.sharepoint.com/drives`,
        accessToken
    )

    for (const drive of drives) {
        let driveNode = {
            id: drive.id,
            name: drive.name,
            children: []
        };

        folderTree.push(driveNode);
        await traverseFolder(driveNode.id, 'root', driveNode.children, accessToken);
    }

    return folderTree
}

async function fetchPdfDownloads(driveId, folderId, accessToken) {
    const url = `https://graph.microsoft.com/v1.0/sites/root/drives/${driveId}/items/${folderId}/children`
    const { value } = await fetchSharePoint(url, accessToken)
  
    if (!value) {
      console.warn('No children found for folder:', folderId)
      return []
    }
    
    // Filter only PDFs and extract name + download URL
    const pdfs = value
      .filter((item) => item.file && item.file.mimeType === "application/pdf")
      .map((item) => ({
        pdfName: item.name,
        downloadUrl: item['@microsoft.graph.downloadUrl']
      }))
  
    return pdfs
  }


// EXPRESS ROUTES

app.post('/accessToken', async (req, res) => {
    const { accessToken } = req.body;
    
    if (!accessToken) {
        return res.status(400).send('Missing access token');
    }

    try {
        const { graphApiToken, refreshToken } = await getTokenOnBehalfOf(accessToken);
        // console.log(`GraphapiToken: ${graphApiToken}, refreshToken: ${refreshToken}`);

        const { graphApiToken: newGraphApiToken, refreshToken: newRefreshToken } = await refreshTokens(refreshToken);
        // console.log(`newGraphapiToken: ${newGraphApiToken}, newRefreshToken: ${newRefreshToken}`);

        const folderTree = await fetchFolderTree(newGraphApiToken);
        console.log("folderTree: ", folderTree);

        const pdfs =  await fetchPdfDownloads('b!gYw8aP-b7EuA6IaFP28xN99zWzd5m-hCh9hq5JYPh2eKb_C17kCvQLcqAhKs8e9f', folderTree[0]?.['children'][0].id, newGraphApiToken)
        console.log("PDFS", pdfs)

        res.status(200).json({ message: 'Access token flow complete', folderTree, pdfs });
    } catch (err) {
        console.error('Error during access token flow:', err);
        res.status(500).send('Server error');
    }
})


// start
const port = process.env.PORT || 3000;
app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});