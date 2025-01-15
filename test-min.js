import * as fs from 'fs'
import { PSTFile } from 'pst-extractor'
import { PSTFolder } from 'pst-extractor'
import { PSTMessage } from 'pst-extractor'

const pstFolder = './testdata/'
const saveToFS = true
const topOutputFolder = './testdataoutput/'

const verbose = true
const displaySender = true
const displayBody = true
let outputFolder = ''
let depth = -1
let col = 0

// console log highlight with https://en.wikipedia.org/wiki/ANSI_escape_code
const ANSI_RED = 31
const ANSI_YELLOW = 93
const ANSI_GREEN = 32
const ANSI_BLUE = 34
const highlight = (str, code = ANSI_RED) => '\u001b[' + code + 'm' + str + '\u001b[0m'

/**
 * Returns a string with visual indication of depth in tree.
 * @param {number} depth
 * @returns {string}
 */
function getDepth(depth) {
  let sdepth = ''
  if (col > 0) {
    col = 0
    sdepth += '\n'
  }
  for (let x = 0; x < depth - 1; x++) {
    sdepth += ' | '
  }
  sdepth += ' |- '
  return sdepth
}

/**
 * Save items to filesystem.
 * @param {PSTMessage} msg
 * @param {string} emailFolder
 * @param {string} sender
 * @param {string} recipients
 */
function doSaveToFS(
  msg,
  emailFolder,
  sender,
  recipients
) {
  try {
    // save the msg as a txt file
    const filename = emailFolder + msg.descriptorNodeId + '.txt'
    if (verbose) {
      console.log(highlight('saving msg to ' + filename))
    }
    const fd = fs.openSync(filename, 'w')
    fs.writeSync(fd, msg.clientSubmitTime + '\r\n')
    fs.writeSync(fd, 'Type: ' + msg.messageClass + '\r\n')
    fs.writeSync(fd, 'From: ' + sender + '\r\n')
    fs.writeSync(fd, 'To: ' + recipients + '\r\n')
    fs.writeSync(fd, 'Subject: ' + msg.subject)
    fs.writeSync(fd, msg.body)
    fs.closeSync(fd)
  } catch (err) {
    console.error(err)
  }

  // walk list of attachments and save to fs
  for (let i = 0; i < msg.numberOfAttachments; i++) {
    const attachment = msg.getAttachment(i)
    // Log.debug1(JSON.stringify(activity, null, 2));
    if (attachment.filename) {
      const filename =
        emailFolder + msg.descriptorNodeId + '-' + attachment.longFilename
      if (verbose) {
        console.log(highlight('saving attachment to ' + filename, ANSI_BLUE))
      }
      try {
        const fd = fs.openSync(filename, 'w')
        const attachmentStream = attachment.fileInputStream
        if (attachmentStream) {
          const bufferSize = 8176
          const buffer = Buffer.alloc(bufferSize)
          let bytesRead
          do {
            bytesRead = attachmentStream.read(buffer)
            fs.writeSync(fd, buffer, 0, bytesRead)
          } while (bytesRead == bufferSize)
          fs.closeSync(fd)
        }
      } catch (err) {
        console.error(err)
      }
    }
  }
}

/**
 * Get the sender and display.
 * @param {PSTMessage} email
 * @returns {string}
 */
function getSender(email) {
  let sender = email.senderName
  if (sender !== email.senderEmailAddress) {
    sender += ' (' + email.senderEmailAddress + ')'
  }
  if (verbose && displaySender && email.messageClass === 'IPM.Note') {
    console.log(getDepth(depth) + ' sender: ' + sender)
  }
  return sender
}

/**
 * Get the recipients and display.
 * @param {PSTMessage} email
 * @returns {string}
 */
function getRecipients(email) {
  // could walk recipients table, but be fast and cheap
  return email.displayTo
}

/**
 * Print a dot representing a message.
 */
function printDot() {
  process.stdout.write('.')
  if (col++ > 100) {
    console.log('')
    col = 0
  }
}

/**
 * Walk the folder tree recursively and process emails.
 * @param {PSTFolder} folder
 */
function processFolder(folder) {
  depth++

  // the root folder doesn't have a display name
  if (depth > 0) {
    console.log(getDepth(depth) + folder.displayName)
  }

  // go through the folders...
  if (folder.hasSubfolders) {
    const childFolders  = folder.getSubFolders()
    for (const childFolder of childFolders) {
      processFolder(childFolder)
    }
  }

  // and now the emails for this folder
  if (folder.contentCount > 0) {
    depth++
    let email = folder.getNextChild()
    while (email != null) {
      if (verbose) {
        console.log(
          getDepth(depth) +
          'Email: ' +
          email.descriptorNodeId +
          ' - ' +
          email.subject
        )
      } else {
        printDot()
      }

      // sender
      const sender = getSender(email)

      // recipients
      const recipients = getRecipients(email)

      // display body?
      if (verbose && displayBody) {
        console.log(highlight('email.body', ANSI_YELLOW), email.body)
        console.log(highlight('email.bodyRTF', ANSI_YELLOW), email.bodyRTF)
        console.log(highlight('email.bodyHTML', ANSI_YELLOW), email.bodyHTML)
      }

      // save content to fs?
      if (saveToFS) {
        // create date string in format YYYY-MM-DD
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

        // create a folder for each day (client submit time)
        const emailFolder = outputFolder + strDate + '/'
        if (!fs.existsSync(emailFolder)) {
          try {
            fs.mkdirSync(emailFolder)
          } catch (err) {
            console.error(err)
          }
        }

        doSaveToFS(email, emailFolder, sender, recipients)
      }
      email = folder.getNextChild()
    }
    depth--
  }
  depth--
}

// make a top level folder to hold content
try {
  if (saveToFS) {
    fs.mkdirSync(topOutputFolder)
  }
} catch (err) {
  console.error(err)
}

const directoryListing = fs.readdirSync(pstFolder)
directoryListing.forEach((filename) => {
  if (filename.endsWith('.pst') || filename.endsWith('.ost')) {
    console.log(highlight(pstFolder + filename, ANSI_GREEN))

    // time for performance comparison to Java and improvement
    const start = Date.now()

    // load file into memory buffer, then open as PSTFile
    const pstFile = new PSTFile(fs.readFileSync(pstFolder + filename))

    // make a sub folder for each PST
    try {
      if (saveToFS) {
        outputFolder = topOutputFolder + filename + '/'
        fs.mkdirSync(outputFolder)
      }
    } catch (err) {
      console.error(err)
    }

    console.log(pstFile.getMessageStore().displayName)
    processFolder(pstFile.getRootFolder())

    const end = Date.now()
    console.log(highlight(pstFolder + filename + ' processed in ' + (end - start) + ' ms', ANSI_GREEN))
  }
})
