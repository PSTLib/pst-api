const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const serverless = require("serverless-http");
const { PSTFile, PSTFolder, PSTMessage } = require('pst-extractor');

const app = express();
app.use(cors()); 
app.use(express.json({ limit: '200mb' }));
app.use(express.urlencoded({ limit: '200mb', extended: true }));
const router = express.Router();

const tempUploadsDir = '/tmp/uploads';
const tempOutputDir = '/tmp/output';

// Ensure temporary directories exist
if (!fs.existsSync(tempUploadsDir)) {
    fs.mkdirSync(tempUploadsDir, { recursive: true });
}
if (!fs.existsSync(tempOutputDir)) {
    fs.mkdirSync(tempOutputDir, { recursive: true });
}


const upload = multer({ 
    dest: tempUploadsDir,
    limits: {
      fileSize: 200 * 1024 * 1024, // 50MB file size limit
    }
});


const processFolder = (zip, outputDir, folder) => {
    let email;
    while (email = folder.getNextChild()) {
        
        let strDate = ''
        let d = email.clientSubmitTime
        if (!d && email.creationTime) {
          d = email.creationTime
        }
        if (d) {
          const month = ('0' + (d.getMonth() + 1)).slice(-2)
          const day = ('0' + d.getDate()).slice(-2)
          strDate = d.getFullYear() + '-' + month + '-' + day
        }
        let body = '';
        if(email.body){
            body = email.body
        } else if(email.bodyRTF){
            body = email.bodyRTF
        } else if(email.bodyHTML){
            body = email.bodyHTML;
        }

        const emailContent = `Subject: ${email.subject}\nFrom: ${email.senderName}\nTo: ${email.displayTo}\nBody: ${body}\n`;

        const emailFileName = `email_${email.descriptorNodeId}.txt`;
        const emailFilePath = path.join(outputDir, emailFileName);
        fs.writeFileSync(emailFilePath, emailContent);
        zip.addLocalFile(emailFilePath, strDate);

        for (let i = 0; i < email.numberOfAttachments; i++) {
            const attachment = email.getAttachment(i);
            if (attachment.filename) {
                const attachmentFileName = `attachment_${email.descriptorNodeId}_${i}-${attachment.longFilename}`;
                const attachmentFilePath = path.join(outputDir, attachmentFileName);
                try {
                    const fd = fs.openSync(attachmentFilePath, 'w');
                    const attachmentStream = attachment.fileInputStream;
                    if (attachmentStream) {
                        const bufferSize = 8176;
                        const buffer = Buffer.alloc(bufferSize);
                        let bytesRead;
                        do {
                            bytesRead = attachmentStream.read(buffer);
                            fs.writeSync(fd, buffer, 0, bytesRead);
                        } while (bytesRead == bufferSize);
                        fs.closeSync(fd);
                    }

                zip.addLocalFile(attachmentFilePath, strDate);
                } catch (err) {
                    console.error(err);
                }
            }
        }
    }
    try{
        folder.getSubFolders().forEach(subFolder => {
            processFolder(zip, outputDir, subFolder);
        });
    }
    catch (err) {
        console.error(err);
    }
};


router.get("/health", (req, res) => {
    res.send("App is running..");
});

router.get("/", (req, res) => {
    res.send("App is running..");
});

router.post('/upload-pst', upload.single('pstFile'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    const pstFilePath = req.file.path;
    const outputDir = path.join(tempOutputDir);
    const zipFileName = `${uuidv4()}.zip`; // Generate a random UUID for the ZIP file
    const zipFilePath = path.join('/tmp', zipFileName);

    try {
        // Extract PST file
        const pstFile = new PSTFile(fs.readFileSync(pstFilePath));
        const zip = new AdmZip();

        processFolder(zip, outputDir, pstFile.getRootFolder());

        // Write ZIP file
        zip.writeZip(zipFilePath);

        res.download(zipFilePath, zipFileName, (err) => {
            if (err) {
                console.error(err);
            }

            // Clean up files
            fs.unlinkSync(pstFilePath);
            fs.unlinkSync(zipFilePath);
        });
    } catch (err) {
        console.error('Error processing PST file:', err);
        res.status(500).send('Error processing PST file.');
    }
    finally {
        // Cleanup
        if (fs.existsSync(pstFilePath)) fs.unlinkSync(pstFilePath);
        if (fs.existsSync(zipFilePath)) fs.unlinkSync(zipFilePath);
    }
});

app.use('/.netlify/functions/app', router);

module.exports = app;
module.exports.handler = serverless(app);