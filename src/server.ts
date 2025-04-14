// src/server.ts
import { FastMCP, UserError } from 'fastmcp';
import { z } from 'zod';
import { google, docs_v1 } from 'googleapis';
import { authorize } from './auth.js';
import { OAuth2Client } from 'google-auth-library';

// --- Helper function for hex color validation (basic) ---
const hexColorRegex = /^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/;
const validateHexColor = (color: string) => hexColorRegex.test(color);

// --- Helper function for Hex to RGB conversion ---
/**
 * Converts a hex color string to a Google Docs API RgbColor object.
 * @param hex - The hex color string (e.g., "#FF0000", "#F00", "FF0000").
 * @returns A Google Docs API RgbColor object or null if invalid.
 */
function hexToRgbColor(hex: string): docs_v1.Schema$RgbColor | null {
  if (!hex) return null;
  let hexClean = hex.startsWith('#') ? hex.slice(1) : hex;

  // Expand shorthand form (e.g. "F00") to full form (e.g. "FF0000")
  if (hexClean.length === 3) {
    hexClean = hexClean[0] + hexClean[0] + hexClean[1] + hexClean[1] + hexClean[2] + hexClean[2];
  }

  if (hexClean.length !== 6) {
    return null; // Invalid length
  }

  const bigint = parseInt(hexClean, 16);
  if (isNaN(bigint)) {
      return null; // Invalid hex characters
  }

  // Extract RGB values and normalize to 0.0 - 1.0 range
  const r = ((bigint >> 16) & 255) / 255;
  const g = ((bigint >> 8) & 255) / 255;
  const b = (bigint & 255) / 255;

  return { red: r, green: g, blue: b };
}

// --- Zod Schema for the formatText tool ---
const FormatTextParameters = z.object({
  documentId: z.string().describe('The ID of the Google Document.'),
  startIndex: z.number().int().min(1).describe('The starting index of the text range (inclusive, starts from 1).'),
  endIndex: z.number().int().min(1).describe('The ending index of the text range (inclusive).'),
  // Optional Formatting Parameters
  bold: z.boolean().optional().describe('Apply bold formatting.'),
  italic: z.boolean().optional().describe('Apply italic formatting.'),
  underline: z.boolean().optional().describe('Apply underline formatting.'),
  strikethrough: z.boolean().optional().describe('Apply strikethrough formatting.'),
  fontSize: z.number().min(1).optional().describe('Set font size (in points, e.g., 12).'),
  fontFamily: z.string().optional().describe('Set font family (e.g., "Arial", "Times New Roman").'),
  foregroundColor: z.string()
    .refine(validateHexColor, { message: "Invalid hex color format (e.g., #FF0000 or #F00)" })
    .optional()
    .describe('Set text color using hex format (e.g., "#FF0000").'),
  backgroundColor: z.string()
    .refine(validateHexColor, { message: "Invalid hex color format (e.g., #00FF00 or #0F0)" })
    .optional()
    .describe('Set text background color using hex format (e.g., "#FFFF00").'),
  linkUrl: z.string().url().optional().describe('Make the text a hyperlink pointing to this URL.')
})
.refine(data => data.endIndex >= data.startIndex, {
  message: "endIndex must be greater than or equal to startIndex",
  path: ["endIndex"], // Point error to endIndex field
})
.refine(data => Object.keys(data).some(key => !['documentId', 'startIndex', 'endIndex'].includes(key) && data[key as keyof typeof data] !== undefined), {
    message: "At least one formatting option (bold, italic, fontSize, etc.) must be provided."
    // No specific path, applies to the whole object
});

// --- Define the TypeScript type based on the schema ---
type FormatTextArgs = z.infer<typeof FormatTextParameters>;

let authClient: OAuth2Client | null = null;
let googleDocs: docs_v1.Docs | null = null;

async function initializeGoogleClient() {
  if (googleDocs) return { authClient, googleDocs };
  if (authClient === null && googleDocs === null) {
    try {
      console.error("Attempting to authorize Google API client...");
      const client = await authorize();
      if (client) {
        authClient = client;
        googleDocs = google.docs({ version: 'v1', auth: authClient });
        console.error("Google API client authorized successfully.");
      } else {
        console.error("FATAL: Authorization returned null or undefined client.");
        authClient = null;
        googleDocs = null;
      }
    } catch (error) {
      console.error("FATAL: Failed to initialize Google API client:", error);
      authClient = null;
      googleDocs = null;
    }
  }
  return { authClient, googleDocs };
}

const server = new FastMCP({
  name: 'Google Docs MCP Server',
  version: '1.0.0',
});

// Tool: Read Google Doc
server.addTool({
  name: 'readGoogleDoc',
  description: 'Reads the content of a specific Google Document.',
  parameters: z.object({
    documentId: z.string().describe('The ID of the Google Document (from the URL).'),
  }),
  execute: async (args, { log }) => {
    const { googleDocs: docs } = await initializeGoogleClient();
    if (!docs) throw new UserError("Google Docs client not initialized.");

    log.info(`Reading Google Doc: ${args.documentId}`);
    try {
      const res = await docs.documents.get({
        documentId: args.documentId,
        fields: 'body(content)',
      });
      log.info(`Fetched doc: ${args.documentId}`);

      let textContent = '';
      res.data.body?.content?.forEach(element => {
        element.paragraph?.elements?.forEach(pe => {
          textContent += pe.textRun?.content || '';
        });
      });

      if (!textContent.trim()) return "Document found, but appears empty.";

      const maxLength = 2000;
      const truncatedContent = textContent.length > maxLength ? textContent.substring(0, maxLength) + '... [truncated]' : textContent;
      return `Content:\n---\n${truncatedContent}`;
    } catch (error: any) {
      log.error(`Error reading doc ${args.documentId}: ${error.message}`);
       if (error.code === 404) throw new UserError(`Doc not found (ID: ${args.documentId}).`);
       if (error.code === 403) throw new UserError(`Permission denied for doc (ID: ${args.documentId}).`);
       throw new UserError(`Failed to read doc: ${error.message}`);
    }
  },
});

// Tool: Append to Google Doc
server.addTool({
  name: 'appendToGoogleDoc',
  description: 'Appends text to the end of a specific Google Document.',
  parameters: z.object({
    documentId: z.string().describe('The ID of the Google Document.'),
    textToAppend: z.string().describe('The text to add.'),
  }),
  execute: async (args, { log }) => {
    const { googleDocs: docs } = await initializeGoogleClient();
     if (!docs) throw new UserError("Google Docs client not initialized.");

    log.info(`Appending to Google Doc: ${args.documentId}`);
    try {
      const docInfo = await docs.documents.get({ documentId: args.documentId, fields: 'body(content)' });
      let endIndex = 1;
      if (docInfo.data.body?.content) {
        const lastElement = docInfo.data.body.content[docInfo.data.body.content.length - 1];
        if (lastElement?.endIndex) endIndex = lastElement.endIndex - 1;
      }
      const textToInsert = (endIndex > 1 && !args.textToAppend.startsWith('\n') ? '\n' : '') + args.textToAppend;

      await docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: { requests: [{ insertText: { location: { index: endIndex }, text: textToInsert } }] },
      });

      log.info(`Successfully appended to doc: ${args.documentId}`);
      return `Successfully appended text to document ${args.documentId}.`;
    } catch (error: any) {
      log.error(`Error editing doc ${args.documentId}: ${error.message}`);
      if (error.code === 404) throw new UserError(`Doc not found (ID: ${args.documentId}).`);
      if (error.code === 403) throw new UserError(`Permission denied for doc (ID: ${args.documentId}).`);
      throw new UserError(`Failed to edit doc: ${error.message}`);
    }
  },
});

// --- Add the NEW formatText tool ---
server.addTool({
  name: 'formatText',
  description: 'Applies character formatting (bold, italics, font size, color, link, etc.) to a specific text range in a Google Document.',
  parameters: FormatTextParameters, // Use the Zod schema defined above
  execute: async (args: FormatTextArgs, { log }) => {
    const { googleDocs: docs } = await initializeGoogleClient();
    if (!docs) {
      throw new UserError("Google Docs client is not initialized. Authentication might have failed.");
    }

    log.info(`Attempting to format text in doc: ${args.documentId}, range: ${args.startIndex}-${args.endIndex}`);

    // 1. Build the TextStyle object based on provided args
    const textStyle: docs_v1.Schema$TextStyle = {};
    const fieldsToUpdate: string[] = []; // To build the 'fields' mask

    if (args.bold !== undefined) { textStyle.bold = args.bold; fieldsToUpdate.push('bold'); }
    if (args.italic !== undefined) { textStyle.italic = args.italic; fieldsToUpdate.push('italic'); }
    if (args.underline !== undefined) { textStyle.underline = args.underline; fieldsToUpdate.push('underline'); }
    if (args.strikethrough !== undefined) { textStyle.strikethrough = args.strikethrough; fieldsToUpdate.push('strikethrough'); }
    if (args.fontSize !== undefined) {
        textStyle.fontSize = { magnitude: args.fontSize, unit: 'PT' };
        fieldsToUpdate.push('fontSize');
    }
    if (args.fontFamily !== undefined) {
        textStyle.weightedFontFamily = { fontFamily: args.fontFamily };
        // Note: API expects 'weightedFontFamily' in fields mask
        fieldsToUpdate.push('weightedFontFamily');
    }
    if (args.foregroundColor !== undefined) {
        const rgbColor = hexToRgbColor(args.foregroundColor);
        if (!rgbColor) throw new UserError(`Invalid foreground hex color format: ${args.foregroundColor}`);
        textStyle.foregroundColor = { color: { rgbColor: rgbColor } };
        fieldsToUpdate.push('foregroundColor');
    }
    if (args.backgroundColor !== undefined) {
        const rgbColor = hexToRgbColor(args.backgroundColor);
        if (!rgbColor) throw new UserError(`Invalid background hex color format: ${args.backgroundColor}`);
        textStyle.backgroundColor = { color: { rgbColor: rgbColor } };
        fieldsToUpdate.push('backgroundColor');
    }
    if (args.linkUrl !== undefined) {
        textStyle.link = { url: args.linkUrl };
        fieldsToUpdate.push('link');
    }

    // Should have already been caught by Zod refine, but double-check
    if (fieldsToUpdate.length === 0) {
        log.warn("No formatting options provided for formatText tool.");
        return "No formatting options were specified.";
    }

    // 2. Build the UpdateTextStyleRequest
    const updateTextStyleRequest: docs_v1.Schema$UpdateTextStyleRequest = {
      range: {
        // API uses segmentId, but omitting it defaults to the document BODY
        startIndex: args.startIndex,
        endIndex: args.endIndex,
        // tabId: TAB_ID, // Optional: specify if working with specific tabs in Sheets embedded in Docs
      },
      textStyle: textStyle,
      fields: fieldsToUpdate.join(','), // Crucial: Tells API which fields to update
    };

    // 3. Build the BatchUpdate request
    const requestBody: docs_v1.Schema$BatchUpdateDocumentRequest = {
      requests: [
        { updateTextStyle: updateTextStyleRequest }
      ]
    };

    // 4. Execute the request
    try {
      await docs.documents.batchUpdate({
        documentId: args.documentId,
        requestBody: requestBody,
      });

      log.info(`Successfully formatted text in doc: ${args.documentId}, range: ${args.startIndex}-${args.endIndex}`);
      return `Successfully applied formatting to range ${args.startIndex}-${args.endIndex} in document ${args.documentId}.`;

    } catch (error: any) {
      log.error(`Error formatting text in doc ${args.documentId}: ${error.message}`, error);
       if (error.code === 404) {
        throw new UserError(`Document not found (ID: ${args.documentId}). Check the ID and permissions.`);
      } else if (error.code === 400 && error.message?.includes('invalid range')) {
         throw new UserError(`Invalid range specified (${args.startIndex}-${args.endIndex}). Check the document length.`);
      } else if (error.code === 403) {
          throw new UserError(`Permission denied for document (ID: ${args.documentId}). Ensure the authenticated user has edit access.`);
      }
      throw new UserError(`Failed to format text: ${error.message || 'Unknown error'}`);
    }
  },
});

// Start the Server (Modified to avoid server.config issue)
async function startServer() {
  await initializeGoogleClient(); // Authorize before starting listeners
  console.error("Starting MCP server...");
  try {
      const configToUse = {
        // Choose one transport:
         transportType: "stdio" as const,
        //  transportType: "sse" as const,
        //  sse: {                       // <-- COMMENT OUT or DELETE SSE config
        //    endpoint: "/sse" as const,
        //    port: 8080,
        //  },
      };

      server.start(configToUse); // Start the server with stdio config

      // Adjust logging (optional, but good practice)
      console.error(`MCP Server running using ${configToUse.transportType}.`);
      if (configToUse.transportType === 'stdio') {
          console.error("Awaiting MCP client connection via stdio...");
      }
      // Removed SSE-specific logging

  } catch(startError) {
      console.error("Error occurred during server.start():", startError);
      throw startError; // Re-throw to be caught by the outer catch
  }
}

// Call the modified startServer function
startServer().catch(err => {
    console.error("Server failed to start:", err);
    process.exit(1);
});