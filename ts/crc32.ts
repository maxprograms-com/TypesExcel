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

export class Crc32 {
    private table: Uint32Array;

    constructor() {
        this.table = this.buildTable();
    }

    compute(data: Buffer): number {
        let crc: number = 0xFFFFFFFF;
        for (let i = 0; i < data.length; i++) {
            crc = this.table[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
        }
        return (crc ^ 0xFFFFFFFF) >>> 0;
    }

    private buildTable(): Uint32Array {
        let table: Uint32Array = new Uint32Array(256);
        for (let n = 0; n < 256; n++) {
            let c: number = n;
            for (let k = 0; k < 8; k++) {
                c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
            }
            table[n] = c >>> 0;
        }
        return table;
    }
}
