import { handleUpload } from '@vercel/blob/client';

const ALLOWED_TYPES = [
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'text/plain',
  'text/markdown',
  'text/csv',
  'application/octet-stream'
];

export default async function handler(request, response) {
  if (request.method !== 'POST') {
    return response.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const body = await request.json();
    const result = await handleUpload({
      body,
      request,
      onBeforeGenerateToken: async (pathname) => ({
        allowedContentTypes: ALLOWED_TYPES,
        maximumSizeInBytes: 50 * 1024 * 1024,
        addRandomSuffix: true,
        tokenPayload: JSON.stringify({ app: 'evalora-ai', purpose: 'project-evaluation' })
      }),
      onUploadCompleted: async ({ blob }) => {
        console.log('Evalora upload completed:', blob.pathname);
      }
    });
    return response.status(200).json(result);
  } catch (error) {
    console.error('Blob upload token error:', error);
    return response.status(400).json({ error: error.message || 'Could not authorize upload' });
  }
}
