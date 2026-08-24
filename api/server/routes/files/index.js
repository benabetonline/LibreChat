const express = require('express');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const db = require('~/models');

const {
  createFileUsageLimiter,
  createFileLimiters,
  configMiddleware,
  requireJwtAuth,
  uaParser,
  checkBan,
} = require('~/server/middleware');
const { restoreTenantContextFromReq } = require('@librechat/api');
const { avatar: asstAvatarRouter } = require('~/server/routes/assistants/v1');
const { avatar: agentAvatarRouter } = require('~/server/routes/agents/v1');
const { createMulterInstance } = require('./multer');

const files = require('./files');
const images = require('./images');
const avatar = require('./avatar');
const speech = require('./speech');

const initialize = async () => {
  const router = express.Router();
  
    router.get('/excel-download/:userId/:fileId', async (req, res) => {
    try {
      const { userId, fileId } = req.params;
      const { expires, signature } = req.query;

      const secret = process.env.JWT_SECRET;

      if (!secret || !expires || !signature) {
        return res.status(403).send('Invalid download link');
      }

      const expiresAt = Number(expires);

      if (!Number.isFinite(expiresAt) || Date.now() > expiresAt) {
        return res.status(403).send('Download link expired');
      }

      if (!/^[a-f0-9]{64}$/i.test(String(signature))) {
        return res.status(403).send('Invalid download signature');
      }

      const payload = `${userId}:${fileId}:${expiresAt}`;

      const expectedSignature = crypto
        .createHmac('sha256', secret)
        .update(payload)
        .digest('hex');

      const receivedBuffer = Buffer.from(String(signature), 'hex');
      const expectedBuffer = Buffer.from(expectedSignature, 'hex');

      if (
        receivedBuffer.length !== expectedBuffer.length ||
        !crypto.timingSafeEqual(receivedBuffer, expectedBuffer)
      ) {
        return res.status(403).send('Invalid download signature');
      }

      const files = await db.getFiles({
        user: userId,
        file_id: fileId,
      });

      const file = files?.[0];

      if (!file || file.source !== 'local') {
        return res.status(404).send('File not found');
      }

      const uploadsRoot = path.resolve('/app/uploads');
      const userRoot = path.resolve(uploadsRoot, String(userId));
      const filePath = path.resolve(file.filepath);

      if (!filePath.startsWith(`${userRoot}${path.sep}`)) {
        return res.status(403).send('Invalid file path');
      }

      if (!fs.existsSync(filePath)) {
        return res.status(404).send('File not found');
      }

      res.setHeader('Cache-Control', 'private, no-store');
      res.setHeader('X-Content-Type-Options', 'nosniff');

      return res.download(
        filePath,
        path.basename(file.filename || 'archivo.xlsx'),
      );
    } catch (error) {
      return res.status(500).send('Error downloading Excel file');
    }
  });
  
  router.use(requireJwtAuth);
  router.use(configMiddleware);
  router.use(checkBan);
  router.use(uaParser);

  const upload = await createMulterInstance();
  router.post('/speech/stt', upload.single('audio'), restoreTenantContextFromReq);

  /* Important: speech route must be added before the upload limiters */
  router.use('/speech', speech);

  const { fileUploadIpLimiter, fileUploadUserLimiter } = createFileLimiters();
  const fileUsageLimiter = createFileUsageLimiter();

  /** Non-strict routing means `/usage/` reaches the same handler, so match the
   *  route the way Express does. An exact comparison would push a
   *  trailing-slash request onto the upload quota instead. */
  const isUsagePath = (path) => path.replace(/\/+$/, '') === '/usage';

  /** Apply rate limiters to all POST routes (excluding /speech which is handled
   *  above). `/usage` is a metadata touch, so it gets its own limiter rather
   *  than consuming upload quota, but it is never left unmetered. */
  router.use((req, res, next) => {
    if (req.method !== 'POST' || req.path.startsWith('/speech')) {
      return next();
    }
    if (isUsagePath(req.path)) {
      return fileUsageLimiter(req, res, next);
    }
    return fileUploadIpLimiter(req, res, (err) => {
      if (err) {
        return next(err);
      }
      return fileUploadUserLimiter(req, res, next);
    });
  });

  router.post('/', upload.single('file'), restoreTenantContextFromReq);
  router.post('/images', upload.single('file'), restoreTenantContextFromReq);
  router.post('/images/avatar', upload.single('file'), restoreTenantContextFromReq);
  router.post(
    '/images/agents/:agent_id/avatar',
    upload.single('file'),
    restoreTenantContextFromReq,
  );
  router.post(
    '/images/assistants/:assistant_id/avatar',
    upload.single('file'),
    restoreTenantContextFromReq,
  );

  router.use('/', files);
  router.use('/images', images);
  router.use('/images/avatar', avatar);
  router.use('/images/agents', agentAvatarRouter);
  router.use('/images/assistants', asstAvatarRouter);
  return router;
};

module.exports = { initialize };
