const path = require('path');
const { config } = require('dotenv');

config({ path: path.join(__dirname, '..', '..', '.env') });

/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  transpilePackages: ['@orcax/api-contracts'],
  poweredByHeader: false,
  env: {
    NEXT_PUBLIC_API_ORIGIN: process.env.NEXT_PUBLIC_API_ORIGIN || 'http://localhost:3001',
    NEXT_PUBLIC_EMAIL_DEV_MAILBOX:
      process.env.EMAIL_TRANSPORT === 'log' && process.env.NODE_ENV !== 'production' ? '1' : '',
  },
};

module.exports = nextConfig;
