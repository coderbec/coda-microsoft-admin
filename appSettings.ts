// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
import * as dotenv from 'dotenv';

// eslint-disable-next-line @typescript-eslint/no-var-requires
require('dotenv').config();

const settings: AppSettings = {
  'clientId': process.env.CLIENT_ID ?? '',
  'clientSecret': process.env.CLIENT_SECRET ?? '',
  'tenantId': process.env.AUTH_TENNANT ?? '',
  'authTenant': process.env.AUTH_TENNANT ?? '',
  'graphUserScopes': [
    'user.read',
    'mail.read',
    'mail.send'
  ]
};

export interface AppSettings {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  authTenant: string;
  graphUserScopes: string[];
}

export default settings;
