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

import { copyFileSync } from 'node:fs';
import { join } from 'node:path';

const assets: string[] = [
    'ts/template.xlsx',
    'i18n/excel_en.json',
    'i18n/excel_es.json',
    'i18n/excel_fr.json'
];

for (let asset of assets) {
    let filename: string = asset.substring(asset.lastIndexOf('/') + 1);
    copyFileSync(asset, join('dist', filename));
}
