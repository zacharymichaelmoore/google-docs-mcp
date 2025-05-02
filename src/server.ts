// src/server.ts
import { FastMCP, UserError } from 'fastmcp';
import { z } from 'zod';
import { google, docs_v1 } from 'googleapis';
import { authorize } from './auth.js';
import { OAuth2Client } from 'google-auth-library';

// Import types and helpers
import {
DocumentIdParameter,
RangeParameters,
OptionalRangeParameters,
TextFindParameter,
TextStyleParameters,
TextStyleArgs,
ParagraphStyleParameters,
ParagraphStyleArgs,
ApplyTextStyleToolParameters, ApplyTextStyleToolArgs,
ApplyParagraphStyleToolParameters, ApplyParagraphStyleToolArgs,
NotImplementedError
} from './types.js';
import * as GDocsHelpers from './googleDocsApiHelpers.js';

let authClient: OAuth2Client | null = null;
let googleDocs: docs_v1.Docs | null = null;

// --- Initialization ---
async function initializeGoogleClient() {
if (googleDocs) return { authClient, googleDocs };
if (!authClient) { // Check authClient instead of googleDocs to allow re-attempt
try {
console.error("Attempting to authorize Google API client...");
const client = await authorize();
authClient = client; // Assign client here
googleDocs = google.docs({ version: 'v1', auth: authClient });
console.error("Google API client authorized successfully.");
} catch (error) {
console.error("FATAL: Failed to initialize Google API client:", error);
authClient = null; // Reset on failure
googleDocs = null;
// Decide if server should exit or just fail tools
throw new Error("Google client initialization failed. Cannot start server tools.");
}
}
// Ensure googleDocs is set if authClient is valid
if (authClient && !googleDocs) {
googleDocs = google.docs({ version: 'v1', auth: authClient });
}

if (!googleDocs) {
throw new Error("Google Docs client could not be initialized.");
}

return { authClient, googleDocs };
}

// Set up process-level unhandled error/rejection handlers to prevent crashes
process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  // Don't exit process, just log the error and continue
  // This will catch timeout errors that might otherwise crash the server
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Promise Rejection:', reason);
  // Don't exit process, just log the error and continue
});

const server = new FastMCP({
  name: 'Ultimate Google Docs MCP Server',
  version: '2.0.0' // Version bump!
});

// --- Helper to get Docs client within tools ---
async function getDocsClient() {
const { googleDocs: docs } = await initializeGoogleClient();
if (!docs) {
throw new UserError("Google Docs client is not initialized. Authentication might have failed during startup or lost connection.");
}
return docs;
}

// === TOOL DEFINITIONS ===

// --- Foundational Tools ---

server.addTool({
name: 'readGoogleDoc',
description: 'Reads the content of a specific Google Document, optionally returning structured data.',
parameters: DocumentIdParameter.extend({
format: z.enum(['text', 'json', 'markdown']).optional().default('text')
.describe("Output format: 'text' (plain text, possibly truncated), 'json' (raw API structure, complex), 'markdown' (experimental conversion).")
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Reading Google Doc: ${args.documentId}, Format: ${args.format}`);

    try {
        const fields = args.format === 'json' || args.format === 'markdown'
            ? '*' // Get everything for structure analysis
            : 'body(content(paragraph(elements(textRun(content)))))'; // Just text content

        const res = await docs.documents.get({
            documentId: args.documentId,
            fields: fields,
        });
        log.info(`Fetched doc: ${args.documentId}`);

        if (args.format === 'json') {
            return JSON.stringify(res.data, null, 2); // Return raw structure
        }

        if (args.format === 'markdown') {
            // TODO: Implement Markdown conversion logic (complex)
            log.warn("Markdown conversion is not implemented yet.");
             throw new NotImplementedError("Markdown output format is not yet implemented.");
            // return convertDocsJsonToMarkdown(res.data);
        }

        // Default: Text format
        let textContent = '';
        res.data.body?.content?.forEach(element => {
            element.paragraph?.elements?.forEach(pe => {
            textContent += pe.textRun?.content || '';
            });
        });

        if (!textContent.trim()) return "Document found, but appears empty.";

        // Basic truncation for text mode
        const maxLength = 4000; // Increased limit
        const truncatedContent = textContent.length > maxLength ? textContent.substring(0, maxLength) + `... [truncated ${textContent.length} chars]` : textContent;
        return `Content:\n---\n${truncatedContent}`;

    } catch (error: any) {
         log.error(`Error reading doc ${args.documentId}: ${error.message || error}`);
         // Handle errors thrown by helpers or API directly
         if (error instanceof UserError) throw error;
         if (error instanceof NotImplementedError) throw error;
         // Generic fallback for API errors not caught by helpers
          if (error.code === 404) throw new UserError(`Doc not found (ID: ${args.documentId}).`);
          if (error.code === 403) throw new UserError(`Permission denied for doc (ID: ${args.documentId}).`);
         throw new UserError(`Failed to read doc: ${error.message || 'Unknown error'}`);
    }

},
});

server.addTool({
name: 'appendToGoogleDoc',
description: 'Appends text to the very end of a specific Google Document.',
parameters: DocumentIdParameter.extend({
textToAppend: z.string().min(1).describe('The text to add to the end.'),
addNewlineIfNeeded: z.boolean().optional().default(true).describe("Automatically add a newline before the appended text if the doc doesn't end with one."),
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Appending to Google Doc: ${args.documentId}`);

    try {
        // Get the current end index
        const docInfo = await docs.documents.get({ documentId: args.documentId, fields: 'body(content(endIndex)),documentStyle(pageSize)' }); // Need content for endIndex
        let endIndex = 1;
        let lastCharIsNewline = false;
        if (docInfo.data.body?.content) {
            const lastElement = docInfo.data.body.content[docInfo.data.body.content.length - 1];
             if (lastElement?.endIndex) {
                endIndex = lastElement.endIndex -1; // Insert *before* the final newline of the doc typically
                // Crude check for last character (better check would involve reading last text run)
                 // const lastTextRun = ... find last text run ...
                 // if (lastTextRun?.content?.endsWith('\n')) lastCharIsNewline = true;
            }
        }
        // Simpler approach: Always assume insertion is needed unless explicitly told not to add newline
        const textToInsert = (args.addNewlineIfNeeded && endIndex > 1 ? '\n' : '') + args.textToAppend;

        if (!textToInsert) return "Nothing to append.";

        const request: docs_v1.Schema$Request = { insertText: { location: { index: endIndex }, text: textToInsert } };
        await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [request]);

        log.info(`Successfully appended to doc: ${args.documentId}`);
        return `Successfully appended text to document ${args.documentId}.`;
    } catch (error: any) {
         log.error(`Error appending to doc ${args.documentId}: ${error.message || error}`);
         if (error instanceof UserError) throw error;
         if (error instanceof NotImplementedError) throw error;
         throw new UserError(`Failed to append to doc: ${error.message || 'Unknown error'}`);
    }

},
});

server.addTool({
name: 'insertText',
description: 'Inserts text at a specific index within the document body.',
parameters: DocumentIdParameter.extend({
textToInsert: z.string().min(1).describe('The text to insert.'),
index: z.number().int().min(1).describe('The index (1-based) where the text should be inserted.'),
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Inserting text in doc ${args.documentId} at index ${args.index}`);
try {
await GDocsHelpers.insertText(docs, args.documentId, args.textToInsert, args.index);
return `Successfully inserted text at index ${args.index}.`;
} catch (error: any) {
log.error(`Error inserting text in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
throw new UserError(`Failed to insert text: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'deleteRange',
description: 'Deletes content within a specified range (start index inclusive, end index exclusive).',
parameters: DocumentIdParameter.extend({
  startIndex: z.number().int().min(1).describe('The starting index of the text range (inclusive, starts from 1).'),
  endIndex: z.number().int().min(1).describe('The ending index of the text range (exclusive).')
}).refine(data => data.endIndex > data.startIndex, {
  message: "endIndex must be greater than startIndex",
  path: ["endIndex"],
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Deleting range ${args.startIndex}-${args.endIndex} in doc ${args.documentId}`);
if (args.endIndex <= args.startIndex) {
throw new UserError("End index must be greater than start index for deletion.");
}
try {
const request: docs_v1.Schema$Request = {
                deleteContentRange: {
                    range: { startIndex: args.startIndex, endIndex: args.endIndex }
                }
            };
            await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [request]);
            return `Successfully deleted content in range ${args.startIndex}-${args.endIndex}.`;
        } catch (error: any) {
            log.error(`Error deleting range in doc ${args.documentId}: ${error.message || error}`);
            if (error instanceof UserError) throw error;
            throw new UserError(`Failed to delete range: ${error.message || 'Unknown error'}`);
}
}
});

// --- Advanced Formatting & Styling Tools ---

server.addTool({
name: 'applyTextStyle',
description: 'Applies character-level formatting (bold, color, font, etc.) to a specific range or found text.',
parameters: ApplyTextStyleToolParameters,
execute: async (args: ApplyTextStyleToolArgs, { log }) => {
const docs = await getDocsClient();
let { startIndex, endIndex } = args.target as any; // Will be updated if target is text

        log.info(`Applying text style in doc ${args.documentId}. Target: ${JSON.stringify(args.target)}, Style: ${JSON.stringify(args.style)}`);

        try {
            // Determine target range
            if ('textToFind' in args.target) {
                const range = await GDocsHelpers.findTextRange(docs, args.documentId, args.target.textToFind, args.target.matchInstance);
                if (!range) {
                    throw new UserError(`Could not find instance ${args.target.matchInstance} of text "${args.target.textToFind}".`);
                }
                startIndex = range.startIndex;
                endIndex = range.endIndex;
                log.info(`Found text "${args.target.textToFind}" (instance ${args.target.matchInstance}) at range ${startIndex}-${endIndex}`);
            }

            if (startIndex === undefined || endIndex === undefined) {
                 throw new UserError("Target range could not be determined.");
            }
             if (endIndex <= startIndex) {
                 throw new UserError("End index must be greater than start index for styling.");
            }

            // Build the request
            const requestInfo = GDocsHelpers.buildUpdateTextStyleRequest(startIndex, endIndex, args.style);
            if (!requestInfo) {
                 return "No valid text styling options were provided.";
            }

            await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [requestInfo.request]);
            return `Successfully applied text style (${requestInfo.fields.join(', ')}) to range ${startIndex}-${endIndex}.`;

        } catch (error: any) {
            log.error(`Error applying text style in doc ${args.documentId}: ${error.message || error}`);
            if (error instanceof UserError) throw error;
            if (error instanceof NotImplementedError) throw error; // Should not happen here
            throw new UserError(`Failed to apply text style: ${error.message || 'Unknown error'}`);
        }
    }

});

server.addTool({
name: 'applyParagraphStyle',
description: 'Applies paragraph-level formatting (alignment, spacing, named styles like Heading 1) to the paragraph(s) containing specific text, an index, or a range.',
parameters: ApplyParagraphStyleToolParameters,
execute: async (args: ApplyParagraphStyleToolArgs, { log }) => {
const docs = await getDocsClient();
let startIndex: number | undefined;
let endIndex: number | undefined;

        log.info(`Applying paragraph style to document ${args.documentId}`);
        log.info(`Style options: ${JSON.stringify(args.style)}`);
        log.info(`Target specification: ${JSON.stringify(args.target)}`);

        try {
            // STEP 1: Determine the target paragraph's range based on the targeting method
            if ('textToFind' in args.target) {
                // Find the text first
                log.info(`Finding text "${args.target.textToFind}" (instance ${args.target.matchInstance || 1})`);
                const textRange = await GDocsHelpers.findTextRange(
                    docs, 
                    args.documentId, 
                    args.target.textToFind, 
                    args.target.matchInstance || 1
                );
                
                if (!textRange) {
                    throw new UserError(`Could not find "${args.target.textToFind}" in the document.`);
                }
                
                log.info(`Found text at range ${textRange.startIndex}-${textRange.endIndex}, now locating containing paragraph`);
                
                // Then find the paragraph containing this text
                const paragraphRange = await GDocsHelpers.getParagraphRange(
                    docs, 
                    args.documentId, 
                    textRange.startIndex
                );
                
                if (!paragraphRange) {
                    throw new UserError(`Found the text but could not determine the paragraph boundaries.`);
                }
                
                startIndex = paragraphRange.startIndex;
                endIndex = paragraphRange.endIndex;
                log.info(`Text is contained within paragraph at range ${startIndex}-${endIndex}`);
                
            } else if ('indexWithinParagraph' in args.target) {
                // Find paragraph containing the specified index
                log.info(`Finding paragraph containing index ${args.target.indexWithinParagraph}`);
                const paragraphRange = await GDocsHelpers.getParagraphRange(
                    docs, 
                    args.documentId, 
                    args.target.indexWithinParagraph
                );
                
                if (!paragraphRange) {
                    throw new UserError(`Could not find paragraph containing index ${args.target.indexWithinParagraph}.`);
                }
                
                startIndex = paragraphRange.startIndex;
                endIndex = paragraphRange.endIndex;
                log.info(`Located paragraph at range ${startIndex}-${endIndex}`);
                
            } else if ('startIndex' in args.target && 'endIndex' in args.target) {
                // Use directly provided range
                startIndex = args.target.startIndex;
                endIndex = args.target.endIndex;
                log.info(`Using provided paragraph range ${startIndex}-${endIndex}`);
            }

            // Verify that we have a valid range
            if (startIndex === undefined || endIndex === undefined) {
                throw new UserError("Could not determine target paragraph range from the provided information.");
            }
            
            if (endIndex <= startIndex) {
                throw new UserError(`Invalid paragraph range: end index (${endIndex}) must be greater than start index (${startIndex}).`);
            }

            // STEP 2: Build and apply the paragraph style request
            log.info(`Building paragraph style request for range ${startIndex}-${endIndex}`);
            const requestInfo = GDocsHelpers.buildUpdateParagraphStyleRequest(startIndex, endIndex, args.style);
            
            if (!requestInfo) {
                return "No valid paragraph styling options were provided.";
            }
            
            log.info(`Applying styles: ${requestInfo.fields.join(', ')}`);
            await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [requestInfo.request]);
            
            return `Successfully applied paragraph styles (${requestInfo.fields.join(', ')}) to the paragraph.`;

        } catch (error: any) {
            // Detailed error logging
            log.error(`Error applying paragraph style in doc ${args.documentId}:`);
            log.error(error.stack || error.message || error);
            
            if (error instanceof UserError) throw error;
            if (error instanceof NotImplementedError) throw error;
            
            // Provide a more helpful error message
            throw new UserError(`Failed to apply paragraph style: ${error.message || 'Unknown error'}`);
        }
    }
});

// --- Structure & Content Tools ---

server.addTool({
name: 'insertTable',
description: 'Inserts a new table with the specified dimensions at a given index.',
parameters: DocumentIdParameter.extend({
rows: z.number().int().min(1).describe('Number of rows for the new table.'),
columns: z.number().int().min(1).describe('Number of columns for the new table.'),
index: z.number().int().min(1).describe('The index (1-based) where the table should be inserted.'),
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Inserting ${args.rows}x${args.columns} table in doc ${args.documentId} at index ${args.index}`);
try {
await GDocsHelpers.createTable(docs, args.documentId, args.rows, args.columns, args.index);
// The API response contains info about the created table, but might be too complex to return here.
return `Successfully inserted a ${args.rows}x${args.columns} table at index ${args.index}.`;
} catch (error: any) {
log.error(`Error inserting table in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
throw new UserError(`Failed to insert table: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'editTableCell',
description: 'Edits the content and/or basic style of a specific table cell. Requires knowing table start index.',
parameters: DocumentIdParameter.extend({
tableStartIndex: z.number().int().min(1).describe("The starting index of the TABLE element itself (tricky to find, may require reading structure first)."),
rowIndex: z.number().int().min(0).describe("Row index (0-based)."),
columnIndex: z.number().int().min(0).describe("Column index (0-based)."),
textContent: z.string().optional().describe("Optional: New text content for the cell. Replaces existing content."),
// Combine basic styles for simplicity here. More advanced cell styling might need separate tools.
textStyle: TextStyleParameters.optional().describe("Optional: Text styles to apply."),
paragraphStyle: ParagraphStyleParameters.optional().describe("Optional: Paragraph styles (like alignment) to apply."),
// cellBackgroundColor: z.string().optional()... // Cell-specific styles are complex
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Editing cell (${args.rowIndex}, ${args.columnIndex}) in table starting at ${args.tableStartIndex}, doc ${args.documentId}`);

        // TODO: Implement complex logic
        // 1. Find the cell's content range based on tableStartIndex, rowIndex, columnIndex. This is NON-TRIVIAL.
        //    Requires getting the document, finding the table element, iterating through rows/cells to calculate indices.
        // 2. If textContent is provided, generate a DeleteContentRange request for the cell's current content.
        // 3. Generate an InsertText request for the new textContent at the cell's start index.
        // 4. If textStyle is provided, generate UpdateTextStyle requests for the new text range.
        // 5. If paragraphStyle is provided, generate UpdateParagraphStyle requests for the cell's paragraph range.
        // 6. Execute batch update.

        log.error("editTableCell is not implemented due to complexity of finding cell indices.");
        throw new NotImplementedError("Editing table cells is complex and not yet implemented.");
        // return `Edit request for cell (${args.rowIndex}, ${args.columnIndex}) submitted (Not Implemented).`;
    }

});

server.addTool({
name: 'insertPageBreak',
description: 'Inserts a page break at the specified index.',
parameters: DocumentIdParameter.extend({
index: z.number().int().min(1).describe('The index (1-based) where the page break should be inserted.'),
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Inserting page break in doc ${args.documentId} at index ${args.index}`);
try {
const request: docs_v1.Schema$Request = {
insertPageBreak: {
location: { index: args.index }
}
};
await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [request]);
return `Successfully inserted page break at index ${args.index}.`;
} catch (error: any) {
log.error(`Error inserting page break in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
throw new UserError(`Failed to insert page break: ${error.message || 'Unknown error'}`);
}
}
});

// --- Intelligent Assistance Tools (Examples/Stubs) ---

server.addTool({
name: 'fixListFormatting',
description: 'EXPERIMENTAL: Attempts to detect paragraphs that look like lists (e.g., starting with -, *, 1.) and convert them to proper Google Docs bulleted or numbered lists. Best used on specific sections.',
parameters: DocumentIdParameter.extend({
// Optional range to limit the scope, otherwise scans whole doc (potentially slow/risky)
range: OptionalRangeParameters.optional().describe("Optional: Limit the fixing process to a specific range.")
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.warn(`Executing EXPERIMENTAL fixListFormatting for doc ${args.documentId}. Range: ${JSON.stringify(args.range)}`);
try {
await GDocsHelpers.detectAndFormatLists(docs, args.documentId, args.range?.startIndex, args.range?.endIndex);
return `Attempted to fix list formatting. Please review the document for accuracy.`;
} catch (error: any) {
log.error(`Error fixing list formatting in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
if (error instanceof NotImplementedError) throw error; // Expected if helper not implemented
throw new UserError(`Failed to fix list formatting: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'addComment',
description: 'Adds a comment anchored to a specific text range. REQUIRES DRIVE API SCOPES/SETUP.',
parameters: DocumentIdParameter.extend({
  startIndex: z.number().int().min(1).describe('The starting index of the text range (inclusive, starts from 1).'),
  endIndex: z.number().int().min(1).describe('The ending index of the text range (exclusive).'),
  commentText: z.string().min(1).describe("The content of the comment."),
}).refine(data => data.endIndex > data.startIndex, {
  message: "endIndex must be greater than startIndex",
  path: ["endIndex"],
}),
execute: async (args, { log }) => {
log.info(`Attempting to add comment "${args.commentText}" to range ${args.startIndex}-${args.endIndex} in doc ${args.documentId}`);
// Requires Drive API client and appropriate scopes.
// const { authClient } = await initializeGoogleClient(); // Get auth client if needed
// if (!authClient) throw new UserError("Authentication client not available for Drive API.");
try {
// await GDocsHelpers.addCommentHelper(driveClient, args.documentId, args.commentText, args.startIndex, args.endIndex);
log.error("addComment requires Drive API setup which is not implemented.");
throw new NotImplementedError("Adding comments requires Drive API setup and is not yet implemented in this server.");
// return `Comment added to range ${args.startIndex}-${args.endIndex}.`;
} catch (error: any) {
log.error(`Error adding comment in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
if (error instanceof NotImplementedError) throw error;
throw new UserError(`Failed to add comment: ${error.message || 'Unknown error'}`);
}
}
});

// --- Add Stubs for other advanced features ---
// (findElement, getDocumentMetadata, replaceText, list management, image handling, section breaks, footnotes, etc.)
// Example Stub:
server.addTool({
name: 'findElement',
description: 'Finds elements (paragraphs, tables, etc.) based on various criteria. (Not Implemented)',
parameters: DocumentIdParameter.extend({
// Define complex query parameters...
textQuery: z.string().optional(),
elementType: z.enum(['paragraph', 'table', 'list', 'image']).optional(),
// styleQuery...
}),
execute: async (args, { log }) => {
log.warn("findElement tool called but is not implemented.");
throw new NotImplementedError("Finding elements by complex criteria is not yet implemented.");
}
});

// --- Preserve the existing formatMatchingText tool for backward compatibility ---
server.addTool({
name: 'formatMatchingText',
description: 'Finds specific text within a Google Document and applies character formatting (bold, italics, color, etc.) to the specified instance.',
parameters: z.object({
  documentId: z.string().describe('The ID of the Google Document.'),
  textToFind: z.string().min(1).describe('The exact text string to find and format.'),
  matchInstance: z.number().int().min(1).optional().default(1).describe('Which instance of the text to format (1st, 2nd, etc.). Defaults to 1.'),
  // Re-use optional Formatting Parameters (SHARED)
  bold: z.boolean().optional().describe('Apply bold formatting.'),
  italic: z.boolean().optional().describe('Apply italic formatting.'),
  underline: z.boolean().optional().describe('Apply underline formatting.'),
  strikethrough: z.boolean().optional().describe('Apply strikethrough formatting.'),
  fontSize: z.number().min(1).optional().describe('Set font size (in points, e.g., 12).'),
  fontFamily: z.string().optional().describe('Set font family (e.g., "Arial", "Times New Roman").'),
  foregroundColor: z.string()
    .refine((color) => /^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/.test(color), { 
      message: "Invalid hex color format (e.g., #FF0000 or #F00)" 
    })
    .optional()
    .describe('Set text color using hex format (e.g., "#FF0000").'),
  backgroundColor: z.string()
    .refine((color) => /^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/.test(color), { 
      message: "Invalid hex color format (e.g., #00FF00 or #0F0)" 
    })
    .optional()
    .describe('Set text background color using hex format (e.g., "#FFFF00").'),
  linkUrl: z.string().url().optional().describe('Make the text a hyperlink pointing to this URL.')
})
.refine(data => Object.keys(data).some(key => !['documentId', 'textToFind', 'matchInstance'].includes(key) && data[key as keyof typeof data] !== undefined), {
    message: "At least one formatting option (bold, italic, fontSize, etc.) must be provided."
}),
execute: async (args, { log }) => {
  // Adapt to use the new applyTextStyle implementation under the hood
  const docs = await getDocsClient();
  log.info(`Using formatMatchingText (legacy) for doc ${args.documentId}, target: "${args.textToFind}" (instance ${args.matchInstance})`);
  
  try {
    // Extract the style parameters
    const styleParams: TextStyleArgs = {};
    if (args.bold !== undefined) styleParams.bold = args.bold;
    if (args.italic !== undefined) styleParams.italic = args.italic;
    if (args.underline !== undefined) styleParams.underline = args.underline;
    if (args.strikethrough !== undefined) styleParams.strikethrough = args.strikethrough;
    if (args.fontSize !== undefined) styleParams.fontSize = args.fontSize;
    if (args.fontFamily !== undefined) styleParams.fontFamily = args.fontFamily;
    if (args.foregroundColor !== undefined) styleParams.foregroundColor = args.foregroundColor;
    if (args.backgroundColor !== undefined) styleParams.backgroundColor = args.backgroundColor;
    if (args.linkUrl !== undefined) styleParams.linkUrl = args.linkUrl;

    // Find the text range
    const range = await GDocsHelpers.findTextRange(docs, args.documentId, args.textToFind, args.matchInstance);
    if (!range) {
      throw new UserError(`Could not find instance ${args.matchInstance} of text "${args.textToFind}".`);
    }
    
    // Build and execute the request
    const requestInfo = GDocsHelpers.buildUpdateTextStyleRequest(range.startIndex, range.endIndex, styleParams);
    if (!requestInfo) {
      return "No valid text styling options were provided.";
    }
    
    await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [requestInfo.request]);
    return `Successfully applied formatting to instance ${args.matchInstance} of "${args.textToFind}".`;
  } catch (error: any) {
    log.error(`Error in formatMatchingText for doc ${args.documentId}: ${error.message || error}`);
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to format text: ${error.message || 'Unknown error'}`);
  }
}
});

// --- Server Startup ---
async function startServer() {
try {
await initializeGoogleClient(); // Authorize BEFORE starting listeners
console.error("Starting Ultimate Google Docs MCP server...");

      // Using stdio as before
      const configToUse = {
          transportType: "stdio" as const,
      };
      
      // Start the server with proper error handling
      server.start(configToUse);
      console.error(`MCP Server running using ${configToUse.transportType}. Awaiting client connection...`);
      
      // Log that error handling has been enabled
      console.error('Process-level error handling configured to prevent crashes from timeout errors.');

} catch(startError: any) {
console.error("FATAL: Server failed to start:", startError.message || startError);
process.exit(1);
}
}

startServer(); // Removed .catch here, let errors propagate if startup fails critically