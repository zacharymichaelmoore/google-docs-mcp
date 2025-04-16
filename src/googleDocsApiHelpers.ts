// src/googleDocsApiHelpers.ts
import { google, docs_v1 } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import { UserError } from 'fastmcp';
import { TextStyleArgs, ParagraphStyleArgs, hexToRgbColor, NotImplementedError } from './types.js';

type Docs = docs_v1.Docs; // Alias for convenience

// --- Constants ---
const MAX_BATCH_UPDATE_REQUESTS = 50; // Google API limits batch size

// --- Core Helper to Execute Batch Updates ---
export async function executeBatchUpdate(docs: Docs, documentId: string, requests: docs_v1.Schema$Request[]): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
if (!requests || requests.length === 0) {
// console.warn("executeBatchUpdate called with no requests.");
return {}; // Nothing to do
}

    // TODO: Consider splitting large request arrays into multiple batches if needed
    if (requests.length > MAX_BATCH_UPDATE_REQUESTS) {
         console.warn(`Attempting batch update with ${requests.length} requests, exceeding typical limits. May fail.`);
    }

    try {
        const response = await docs.documents.batchUpdate({
            documentId: documentId,
            requestBody: { requests },
        });
        return response.data;
    } catch (error: any) {
        console.error(`Google API batchUpdate Error for doc ${documentId}:`, error.response?.data || error.message);
        // Translate common API errors to UserErrors
        if (error.code === 400 && error.message.includes('Invalid requests')) {
             // Try to extract more specific info if available
             const details = error.response?.data?.error?.details;
             let detailMsg = '';
             if (details && Array.isArray(details)) {
                 detailMsg = details.map(d => d.description || JSON.stringify(d)).join('; ');
             }
            throw new UserError(`Invalid request sent to Google Docs API. Details: ${detailMsg || error.message}`);
        }
        if (error.code === 404) throw new UserError(`Document not found (ID: ${documentId}). Check the ID.`);
        if (error.code === 403) throw new UserError(`Permission denied for document (ID: ${documentId}). Ensure the authenticated user has edit access.`);
        // Generic internal error for others
        throw new Error(`Google API Error (${error.code}): ${error.message}`);
    }

}

// --- Text Finding Helper ---
// NOTE: This is a simplified version. A robust version needs to handle
// text spanning multiple TextRuns, pagination, tables etc.
export async function findTextRange(docs: Docs, documentId: string, textToFind: string, instance: number = 1): Promise<{ startIndex: number; endIndex: number } | null> {
try {
const res = await docs.documents.get({
documentId,
fields: 'body(content(paragraph(elements(startIndex,endIndex,textRun(content)))))',
});

        if (!res.data.body?.content) return null;

        let fullText = '';
        const segments: { text: string, start: number, end: number }[] = [];
        res.data.body.content.forEach(element => {
            element.paragraph?.elements?.forEach(pe => {
                if (pe.textRun?.content && pe.startIndex && pe.endIndex) {
                    const content = pe.textRun.content;
                    fullText += content;
                    segments.push({ text: content, start: pe.startIndex, end: pe.endIndex });
                }
            });
        });

        let startIndex = -1;
        let endIndex = -1;
        let foundCount = 0;
        let searchStartIndex = 0;

        while (foundCount < instance) {
            const currentIndex = fullText.indexOf(textToFind, searchStartIndex);
            if (currentIndex === -1) break;

            foundCount++;
            if (foundCount === instance) {
                const targetStartInFullText = currentIndex;
                const targetEndInFullText = currentIndex + textToFind.length;
                let currentPosInFullText = 0;

                for (const seg of segments) {
                    const segStartInFullText = currentPosInFullText;
                    const segTextLength = seg.text.length;
                    const segEndInFullText = segStartInFullText + segTextLength;

                    if (startIndex === -1 && targetStartInFullText >= segStartInFullText && targetStartInFullText < segEndInFullText) {
                        startIndex = seg.start + (targetStartInFullText - segStartInFullText);
                    }
                    if (targetEndInFullText > segStartInFullText && targetEndInFullText <= segEndInFullText) {
                        endIndex = seg.start + (targetEndInFullText - segStartInFullText);
                        break;
                    }
                    currentPosInFullText = segEndInFullText;
                }

                if (startIndex === -1 || endIndex === -1) { // Mapping failed for this instance
                    startIndex = -1; endIndex = -1;
                    // Continue searching from *after* this failed mapping attempt
                    searchStartIndex = currentIndex + 1;
                    foundCount--; // Decrement count as this instance wasn't successfully mapped
                    continue;
                }
                // Successfully mapped
                 return { startIndex, endIndex };
            }
             // Prepare for next search iteration
            searchStartIndex = currentIndex + 1;
        }

        return null; // Instance not found or mapping failed for all attempts
    } catch (error: any) {
        console.error(`Error finding text "${textToFind}" in doc ${documentId}: ${error.message}`);
         if (error.code === 404) throw new UserError(`Document not found while searching text (ID: ${documentId}).`);
         if (error.code === 403) throw new UserError(`Permission denied while searching text in doc (ID: ${documentId}).`);
        throw new Error(`Failed to retrieve doc for text searching: ${error.message}`);
    }

}

// --- Paragraph Boundary Helper ---
// Finds the paragraph containing a given index. Very simplified.
// A robust version needs to understand structural elements better.
export async function getParagraphRange(docs: Docs, documentId: string, indexWithin: number): Promise<{ startIndex: number; endIndex: number } | null> {
try {
const res = await docs.documents.get({
documentId,
// Request paragraph elements and their ranges
fields: 'body(content(startIndex,endIndex,paragraph))',
});

        if (!res.data.body?.content) return null;

        for (const element of res.data.body.content) {
            if (element.paragraph && element.startIndex && element.endIndex) {
                // Check if the provided index falls within this paragraph element's range
                // API ranges are typically [startIndex, endIndex)
                if (indexWithin >= element.startIndex && indexWithin < element.endIndex) {
                    return { startIndex: element.startIndex, endIndex: element.endIndex };
                }
            }
        }
        return null; // Index not found within any paragraph element

    } catch (error: any) {
        console.error(`Error getting paragraph range for index ${indexWithin} in doc ${documentId}: ${error.message}`);
         if (error.code === 404) throw new UserError(`Document not found while finding paragraph range (ID: ${documentId}).`);
         if (error.code === 403) throw new UserError(`Permission denied while finding paragraph range in doc (ID: ${documentId}).`);
        throw new Error(`Failed to retrieve doc for paragraph range finding: ${error.message}`);
    }

}

// --- Style Request Builders ---

export function buildUpdateTextStyleRequest(
startIndex: number,
endIndex: number,
style: TextStyleArgs
): { request: docs_v1.Schema$Request, fields: string[] } | null {
    const textStyle: docs_v1.Schema$TextStyle = {};
const fieldsToUpdate: string[] = [];

    if (style.bold !== undefined) { textStyle.bold = style.bold; fieldsToUpdate.push('bold'); }
    if (style.italic !== undefined) { textStyle.italic = style.italic; fieldsToUpdate.push('italic'); }
    if (style.underline !== undefined) { textStyle.underline = style.underline; fieldsToUpdate.push('underline'); }
    if (style.strikethrough !== undefined) { textStyle.strikethrough = style.strikethrough; fieldsToUpdate.push('strikethrough'); }
    if (style.fontSize !== undefined) { textStyle.fontSize = { magnitude: style.fontSize, unit: 'PT' }; fieldsToUpdate.push('fontSize'); }
    if (style.fontFamily !== undefined) { textStyle.weightedFontFamily = { fontFamily: style.fontFamily }; fieldsToUpdate.push('weightedFontFamily'); }
    if (style.foregroundColor !== undefined) {
        const rgbColor = hexToRgbColor(style.foregroundColor);
        if (!rgbColor) throw new UserError(`Invalid foreground hex color format: ${style.foregroundColor}`);
        textStyle.foregroundColor = { color: { rgbColor: rgbColor } }; fieldsToUpdate.push('foregroundColor');
    }
     if (style.backgroundColor !== undefined) {
        const rgbColor = hexToRgbColor(style.backgroundColor);
        if (!rgbColor) throw new UserError(`Invalid background hex color format: ${style.backgroundColor}`);
        textStyle.backgroundColor = { color: { rgbColor: rgbColor } }; fieldsToUpdate.push('backgroundColor');
    }
    if (style.linkUrl !== undefined) {
        textStyle.link = { url: style.linkUrl }; fieldsToUpdate.push('link');
    }
    // TODO: Handle clearing formatting

    if (fieldsToUpdate.length === 0) return null; // No styles to apply

    const request: docs_v1.Schema$Request = {
        updateTextStyle: {
            range: { startIndex, endIndex },
            textStyle: textStyle,
            fields: fieldsToUpdate.join(','),
        }
    };
    return { request, fields: fieldsToUpdate };

}

export function buildUpdateParagraphStyleRequest(
startIndex: number,
endIndex: number,
style: ParagraphStyleArgs
): { request: docs_v1.Schema$Request, fields: string[] } | null {
    const paragraphStyle: docs_v1.Schema$ParagraphStyle = {};
const fieldsToUpdate: string[] = [];

    if (style.alignment !== undefined) { paragraphStyle.alignment = style.alignment; fieldsToUpdate.push('alignment'); }
    if (style.indentStart !== undefined) { paragraphStyle.indentStart = { magnitude: style.indentStart, unit: 'PT' }; fieldsToUpdate.push('indentStart'); }
    if (style.indentEnd !== undefined) { paragraphStyle.indentEnd = { magnitude: style.indentEnd, unit: 'PT' }; fieldsToUpdate.push('indentEnd'); }
    if (style.spaceAbove !== undefined) { paragraphStyle.spaceAbove = { magnitude: style.spaceAbove, unit: 'PT' }; fieldsToUpdate.push('spaceAbove'); }
    if (style.spaceBelow !== undefined) { paragraphStyle.spaceBelow = { magnitude: style.spaceBelow, unit: 'PT' }; fieldsToUpdate.push('spaceBelow'); }
    if (style.namedStyleType !== undefined) { paragraphStyle.namedStyleType = style.namedStyleType; fieldsToUpdate.push('namedStyleType'); }
    if (style.keepWithNext !== undefined) { paragraphStyle.keepWithNext = style.keepWithNext; fieldsToUpdate.push('keepWithNext'); }
     // TODO: Handle borders, clearing formatting

    if (fieldsToUpdate.length === 0) return null; // No styles to apply

     const request: docs_v1.Schema$Request = {
        updateParagraphStyle: {
            range: { startIndex, endIndex },
            paragraphStyle: paragraphStyle,
            fields: fieldsToUpdate.join(','),
        }
    };
     return { request, fields: fieldsToUpdate };

}

// --- Specific Feature Helpers ---

export async function createTable(docs: Docs, documentId: string, rows: number, columns: number, index: number): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    if (rows < 1 || columns < 1) {
        throw new UserError("Table must have at least 1 row and 1 column.");
    }
    const request: docs_v1.Schema$Request = {
insertTable: {
location: { index },
rows: rows,
columns: columns,
}
};
return executeBatchUpdate(docs, documentId, [request]);
}

export async function insertText(docs: Docs, documentId: string, text: string, index: number): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    if (!text) return {}; // Nothing to insert
    const request: docs_v1.Schema$Request = {
insertText: {
location: { index },
text: text,
}
};
return executeBatchUpdate(docs, documentId, [request]);
}

// --- Complex / Stubbed Helpers ---

export async function findParagraphsMatchingStyle(
docs: Docs,
documentId: string,
styleCriteria: any // Define a proper type for criteria (e.g., { fontFamily: 'Arial', bold: true })
): Promise<{ startIndex: number; endIndex: number }[]> {
// TODO: Implement logic
// 1. Get document content with paragraph elements and their styles.
// 2. Iterate through paragraphs.
// 3. For each paragraph, check if its computed style matches the criteria.
// 4. Return ranges of matching paragraphs.
console.warn("findParagraphsMatchingStyle is not implemented.");
throw new NotImplementedError("Finding paragraphs by style criteria is not yet implemented.");
// return [];
}

export async function detectAndFormatLists(
docs: Docs,
documentId: string,
startIndex?: number,
endIndex?: number
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
// TODO: Implement complex logic
// 1. Get document content (paragraphs, text runs) in the specified range (or whole doc).
// 2. Iterate through paragraphs.
// 3. Identify sequences of paragraphs starting with list-like markers (e.g., "-", "*", "1.", "a)").
// 4. Determine nesting levels based on indentation or marker patterns.
// 5. Generate CreateParagraphBulletsRequests for the identified sequences.
// 6. Potentially delete the original marker text.
// 7. Execute the batch update.
console.warn("detectAndFormatLists is not implemented.");
throw new NotImplementedError("Automatic list detection and formatting is not yet implemented.");
// return {};
}

export async function addCommentHelper(docs: Docs, documentId: string, text: string, startIndex: number, endIndex: number): Promise<void> {
// NOTE: Adding comments typically requires the Google Drive API v3 and different scopes!
// 'https://www.googleapis.com/auth/drive' or more specific comment scopes.
// This helper is a placeholder assuming Drive API client (`drive`) is available and authorized.
/*
const drive = google.drive({version: 'v3', auth: authClient}); // Assuming authClient is available
await drive.comments.create({
fileId: documentId,
requestBody: {
content: text,
anchor: JSON.stringify({ // Anchor format might need verification
'type': 'workbook#textAnchor', // Or appropriate type for Docs
'refs': [{
'docRevisionId': 'head', // Or specific revision
'range': {
'start': startIndex,
'end': endIndex,
}
}]
})
},
fields: 'id'
});
*/
console.warn("addCommentHelper requires Google Drive API and is not implemented.");
throw new NotImplementedError("Adding comments requires Drive API setup and is not yet implemented.");
}