/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { Dexie, Table } from 'dexie';
import { Session, Folder } from './types';

export class ResultDatabase extends Dexie {
  sessions!: Table<Session>;
  folders!: Table<Folder>;

  constructor() {
    super('ResultManagerDB');
    this.version(2).stores({
      sessions: 'id, name, createdAt, folderId', // Primary key and indexed props
      folders: 'id, name, createdAt'
    });
  }
}

export const db = new ResultDatabase();