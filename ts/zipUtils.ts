/*******************************************************************************
 * Copyright (c) 2026 Maxprograms.
 *
 * This program and the accompanying materials
 * are made available under the terms of the Eclipse Public License 1.0
 * which accompanies this distribution, and is available at
 * https://www.eclipse.org/org/documents/epl-v10.html
 *
 * Contributors:
 *     Maxprograms - initial API and implementation
 *******************************************************************************/

import { deflateRawSync, inflateRawSync } from 'node:zlib';
import { Crc32 } from './crc32.js';

const LOCAL_FILE_HEADER_SIGNATURE: number = 0x04034b50;
const CENTRAL_DIRECTORY_SIGNATURE: number = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE: number = 0x06054b50;
const METHOD_STORED: number = 0;
const METHOD_DEFLATE: number = 8;

interface CentralDirectoryEntry {
    name: string;
    method: number;
    compressedSize: number;
    uncompressedSize: number;
    localHeaderOffset: number;
}

export class ZipUtils {
    private crc32: Crc32;

    constructor() {
        this.crc32 = new Crc32();
    }

    readZip(buffer: Buffer): Map<string, Buffer> {
        let eocdOffset: number = this.findEndOfCentralDirectory(buffer);
        let totalEntries: number = buffer.readUInt16LE(eocdOffset + 10);
        let centralDirectoryOffset: number = buffer.readUInt32LE(eocdOffset + 16);

        let entries: CentralDirectoryEntry[] = [];
        let pointer: number = centralDirectoryOffset;
        for (let i = 0; i < totalEntries; i++) {
            if (buffer.readUInt32LE(pointer) !== CENTRAL_DIRECTORY_SIGNATURE) {
                throw new Error('Corrupt zip entry: central directory signature mismatch');
            }
            let method: number = buffer.readUInt16LE(pointer + 10);
            let compressedSize: number = buffer.readUInt32LE(pointer + 20);
            let uncompressedSize: number = buffer.readUInt32LE(pointer + 24);
            let nameLength: number = buffer.readUInt16LE(pointer + 28);
            let extraLength: number = buffer.readUInt16LE(pointer + 30);
            let commentLength: number = buffer.readUInt16LE(pointer + 32);
            let localHeaderOffset: number = buffer.readUInt32LE(pointer + 42);
            let name: string = buffer.toString('utf8', pointer + 46, pointer + 46 + nameLength);

            entries.push({ name, method, compressedSize, uncompressedSize, localHeaderOffset });
            pointer += 46 + nameLength + extraLength + commentLength;
        }

        let result: Map<string, Buffer> = new Map();
        for (let entry of entries) {
            result.set(entry.name, this.extractEntry(buffer, entry));
        }
        return result;
    }

    writeZip(entries: Map<string, Buffer>): Buffer {
        let localParts: Buffer[] = [];
        let centralParts: Buffer[] = [];
        let offset: number = 0;
        let { time, date } = this.dosDateTime(new Date());

        for (let [name, data] of entries) {
            let nameBuffer: Buffer = Buffer.from(name, 'utf8');
            let compressed: Buffer = deflateRawSync(data);
            let checksum: number = this.crc32.compute(data);

            let localHeader: Buffer = Buffer.alloc(30);
            localHeader.writeUInt32LE(LOCAL_FILE_HEADER_SIGNATURE, 0);
            localHeader.writeUInt16LE(20, 4);
            localHeader.writeUInt16LE(0, 6);
            localHeader.writeUInt16LE(METHOD_DEFLATE, 8);
            localHeader.writeUInt16LE(time, 10);
            localHeader.writeUInt16LE(date, 12);
            localHeader.writeUInt32LE(checksum, 14);
            localHeader.writeUInt32LE(compressed.length, 18);
            localHeader.writeUInt32LE(data.length, 22);
            localHeader.writeUInt16LE(nameBuffer.length, 26);
            localHeader.writeUInt16LE(0, 28);

            localParts.push(localHeader, nameBuffer, compressed);

            let centralHeader: Buffer = Buffer.alloc(46);
            centralHeader.writeUInt32LE(CENTRAL_DIRECTORY_SIGNATURE, 0);
            centralHeader.writeUInt16LE(20, 4);
            centralHeader.writeUInt16LE(20, 6);
            centralHeader.writeUInt16LE(0, 8);
            centralHeader.writeUInt16LE(METHOD_DEFLATE, 10);
            centralHeader.writeUInt16LE(time, 12);
            centralHeader.writeUInt16LE(date, 14);
            centralHeader.writeUInt32LE(checksum, 16);
            centralHeader.writeUInt32LE(compressed.length, 20);
            centralHeader.writeUInt32LE(data.length, 24);
            centralHeader.writeUInt16LE(nameBuffer.length, 28);
            centralHeader.writeUInt16LE(0, 30);
            centralHeader.writeUInt16LE(0, 32);
            centralHeader.writeUInt16LE(0, 34);
            centralHeader.writeUInt16LE(0, 36);
            centralHeader.writeUInt32LE(0, 38);
            centralHeader.writeUInt32LE(offset, 42);

            centralParts.push(centralHeader, nameBuffer);

            offset += localHeader.length + nameBuffer.length + compressed.length;
        }

        let centralDirectory: Buffer = Buffer.concat(centralParts);
        let centralDirectoryOffset: number = offset;

        let eocd: Buffer = Buffer.alloc(22);
        eocd.writeUInt32LE(END_OF_CENTRAL_DIRECTORY_SIGNATURE, 0);
        eocd.writeUInt16LE(0, 4);
        eocd.writeUInt16LE(0, 6);
        eocd.writeUInt16LE(entries.size, 8);
        eocd.writeUInt16LE(entries.size, 10);
        eocd.writeUInt32LE(centralDirectory.length, 12);
        eocd.writeUInt32LE(centralDirectoryOffset, 16);
        eocd.writeUInt16LE(0, 20);

        return Buffer.concat([...localParts, centralDirectory, eocd]);
    }

    private extractEntry(buffer: Buffer, entry: CentralDirectoryEntry): Buffer {
        let offset: number = entry.localHeaderOffset;
        if (buffer.readUInt32LE(offset) !== LOCAL_FILE_HEADER_SIGNATURE) {
            throw new Error('Corrupt zip entry: local file header signature mismatch for ' + entry.name);
        }
        let nameLength: number = buffer.readUInt16LE(offset + 26);
        let extraLength: number = buffer.readUInt16LE(offset + 28);
        let dataStart: number = offset + 30 + nameLength + extraLength;
        let compressed: Buffer = buffer.subarray(dataStart, dataStart + entry.compressedSize);

        if (entry.method === METHOD_STORED) {
            return Buffer.from(compressed);
        }
        if (entry.method === METHOD_DEFLATE) {
            return inflateRawSync(compressed);
        }
        throw new Error('Unsupported compression method (' + entry.method + ') for ' + entry.name);
    }

    private findEndOfCentralDirectory(buffer: Buffer): number {
        let minOffset: number = Math.max(0, buffer.length - 22 - 0xFFFF);
        for (let offset = buffer.length - 22; offset >= minOffset; offset--) {
            if (buffer.readUInt32LE(offset) === END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
                return offset;
            }
        }
        throw new Error('Not a valid zip file: end of central directory record not found');
    }

    private dosDateTime(value: Date): { time: number, date: number } {
        let year: number = Math.max(1980, value.getFullYear());
        let date: number = ((year - 1980) << 9) | ((value.getMonth() + 1) << 5) | value.getDate();
        let time: number = (value.getHours() << 11) | (value.getMinutes() << 5) | Math.floor(value.getSeconds() / 2);
        return { time, date };
    }
}
