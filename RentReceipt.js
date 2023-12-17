/**
 * Script written by Md Arifuzzzaman
 * Date: 12/13/2023
 * Description: This script automates the Rent Receipt process in Google Sheets and can accomplish the following tasks:
 *   1. Record and display payment information received via Cash App in a designated sheet.
 *   2. Generate and rename a PDF receipt for each payment.
 *   3. Manage email communication and create drafts for sending receipts.
 *   4. Calculate and display late fees in cell F43 based on payment and rent amounts (if applicable).
 *   5. Handle email subject extraction and dollar amount recognition.
 *   6. Provide a custom menu option to run the script within Google Sheets.
 */

function RentReceipt() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Rent_Receipt');

    if (!sheet) {
        throw new Error('Sheet not found. Check the sheet name for typos.');
    }

    var tenantName = sheet.getRange('D9').getValue();
    var emailDetails = getLatestCashAppEmailSubjectAndDate(tenantName);

    if (emailDetails) {
        var emailReceivedInfo = "Payment Received on " + Utilities.formatDate(emailDetails.date, Session.getScriptTimeZone(), "MMM d, yyyy, h:mm a") + " - " + emailDetails.subject;
        sheet.getRange('B21').setValue(emailReceivedInfo);

        sheet.getRange('F4').setValue(Utilities.formatDate(emailDetails.date, Session.getScriptTimeZone(), "MM/dd/yyyy"));
        var currentMonth = emailDetails.date.getMonth() + 1;
        sheet.getRange('F6').setValue(`RECEIPT NO: ${currentMonth}`);

        var dollarAmount = extractDollarAmount(emailDetails.subject);
        if (dollarAmount) {
            sheet.getRange('F21').setValue(dollarAmount);
        } else {
            sheet.getRange('F21').clearContent();
        }

        var newTitle = "2446 Ontario Ave_APT#3_" + Utilities.formatDate(emailDetails.date, Session.getScriptTimeZone(), "MMMM_yyyy") + "_Rent_Receipt";
        spreadsheet.rename(newTitle);

        // Check if the PDF exists
        var pdfFiles = DriveApp.getFilesByName(newTitle + ".pdf");

        while (pdfFiles.hasNext()) {
            var pdf = pdfFiles.next();
            pdf.setTrashed(true); // Delete existing PDF
        }

        // Create a new PDF
        var blob = spreadsheet.getBlob();
        blob.setName(newTitle + ".pdf");
        var pdf = DriveApp.createFile(blob);

        var emailSubjectMd = getLatestEmailSubject('md@clevvarestate.com');

        if (emailSubjectMd) {
            var existingSubjects = sheet.getRange('B21:B33').getValues();
            var targetCell;

            // Check if B21 is empty or has a value
            if (!existingSubjects[0][0]) {
                targetCell = sheet.getRange('B21');
            } else if (!existingSubjects[4][0]) {
                targetCell = sheet.getRange('B25');
            } else if (!existingSubjects[8][0]) {
                targetCell = sheet.getRange('B29');
            } else if (!existingSubjects[12][0]) {
                targetCell = sheet.getRange('B33');
            } else {
                console.log('All target cells (B21, B25, B29, B33) are filled.');
            }

            if (targetCell) {
                // Check if the subject is not already in the target cells to avoid duplication
                var isDuplicate = false;
                for (var i = 0; i < existingSubjects.length; i++) {
                    if (existingSubjects[i][0] === emailSubjectMd) {
                        isDuplicate = true;
                        break;
                    }
                }

                if (!isDuplicate) {
                    targetCell.setValue(emailSubjectMd);

                    var dollarAmountFromSubject = extractDollarAmount(emailSubjectMd);
                    if (dollarAmountFromSubject) {
                        sheet.getRange('F' + (targetCell.getRow())).setValue(dollarAmountFromSubject);
                    } else {
                        sheet.getRange('F' + (targetCell.getRow())).clearContent();
                    }
                } else {
                    console.log('Email subject is a duplicate and will not be inserted.');
                }
            }
        } else {
            console.log('No recent email found from "md@clevvarestate.com".');
        }

        // Check if F42 is greater than or equal to C5, then calculate late fee and add to F43
        var rentAmount = sheet.getRange('C5').getValue();
        var currentTotal = sheet.getRange('F42').getValue();
        var lateFeeCell = sheet.getRange('F43');

        if (currentTotal > rentAmount) {
            var lateFee = currentTotal - rentAmount;
            lateFeeCell.setValue(lateFee);
        } else {
            lateFeeCell.clearContent(); // Clear F43 if no late fee
        }

        // Compose the email draft with tenant's email in the "to" field
        var tenantEmail = sheet.getRange('D12').getValue();
        var ccEmail = sheet.getRange('B14').getValue();
        var emailBodyWithSignature = emailReceivedInfo + getEmailSignature();
        GmailApp.createDraft(tenantEmail, emailDetails.subject, emailBodyWithSignature, {
            attachments: [pdf],
            cc: ccEmail
        });
    } else {
        console.log('No recent Cash App email found for the tenant.');
    }
}

function getEmailSignature() {
    var emailSignature = "\n\nBest Regards, \n\nProperty Management Team \n\nClevvar Estate\nCall or Text: (716) 236-8207\nadmin@clevvarestate.com\nOfficial Website: https://clevvarestate.com\n\nThis email, including any attachments, may contain confidential, privileged, or otherwise legally protected information intended solely for the person(s) or entity(ies) to which it is addressed. If you are not the intended recipient, you are hereby notified that any dissemination, distribution, copying, or other use of the email or its attachments is prohibited. Please immediately notify the sender of your access to the email or its attachments by replying to the message and delete all copies.";
    return emailSignature;
}

function getLatestCashAppEmailSubjectAndDate(tenantName) {
    var threads = GmailApp.search(`from:cash@square.com "${tenantName}"`, 0, 1);

    if (threads.length > 0) {
        var message = threads[0].getMessages()[0];
        var subject = message.getSubject();
        var date = message.getDate();

        if (subject.includes(tenantName)) {
            return { subject: subject, date: date };
        }
    }
    return null;
}

function getLatestEmailSubject(sender) {
    var threads = GmailApp.search(`from:${sender}`, 0, 1);

    if (threads.length > 0) {
        var message = threads[0].getMessages()[0];
        var subject = message.getSubject();
        return subject;
    }
    return null;
}

function extractDollarAmount(subject) {
    var amountPattern = /\$\d+(?:\.\d{1,2})?/;
    var matches = subject.match(amountPattern;
    return matches ? matches[0] : null;
}

function setupForTesting() {
    RentReceipt();
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Code')
        .addItem('Run Rent Receipt Script', 'RentReceipt')
        .addToUi();
}
