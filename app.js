const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const { PSTFile, PSTFolder, PSTMessage } = require('pst-extractor');

const app = express();
app.use(cors()); 
app.use(express.json({ limit: '200mb' }));
app.use(express.urlencoded({ limit: '200mb', extended: true }));

const upload = multer({ dest: 'uploads/',
    limits: {
      fileSize: 200 * 1024 * 1024, // 50MB file size limit
    } });

// Ensure uploads and output directories exist
if (!fs.existsSync('uploads')) {
    fs.mkdirSync('uploads');
}
if (!fs.existsSync('output')) {
    fs.mkdirSync('output');
}

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
        }
        else if(email.bodyRTF){
            body = email.bodyRTF
        }
        else if(email.bodyHTML){
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

app.post('/upload-pst', upload.single('pstFile'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    const pstFilePath = req.file.path;
    const outputDir = path.join(__dirname, 'output');
    const zipFileName = `${uuidv4()}.zip`; // Generate a random UUID for the ZIP file
    const zipFilePath = path.join(__dirname, zipFileName);

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
});

app.listen(3000, () => {
    console.log('Server started on port 3000');
});