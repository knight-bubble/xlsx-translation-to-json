import { translate } from 'free-translate';
import { Locale } from 'free-translate/dist/types/locales';
import { writeFileSync } from 'fs';
import * as reader from 'xlsx';

interface Translation {
  Label: string;
  English: string;
  France: string;
  Spanish: string;
  Danish: string;
}

const langs = {
  Spanish: 'es',
  English: 'en',
  France: 'fr',
  Danish: 'da',
};

const file = reader.readFile('./Javelin Translation - Master Sheet.xlsx');
const reference = require('./reference.json');

const sheetName = 'Work order';
const language = 'Spanish';
const dictName = 'WORKORDER';

let data: Record<string, Record<string, string>> = { [dictName]: {} };
const temp = reader.utils.sheet_to_json(file.Sheets[sheetName]) as Translation[];
temp.forEach((translation) => {
  let [dict, label] = translation.Label.split('/');
  if (dict.indexOf(':') > 0) {
    label = dict.split(':')[0];
  }
  label = label.replace(/[\"\s]/g, '');
  data[dictName][label] = translation[language];
});

const keysPresent = Object.keys(data[dictName]);
const refKeys = Object.keys(reference[dictName]);

let missingKeys = 0;
const missingTranslationsForXlsx: any[] = [];
const missingTranslations: typeof data = { [dictName]: {} };
async function translateMissing() {
  for (const rKey of refKeys) {
    if (keysPresent.indexOf(rKey) > 0) {
      continue;
    }

    const translatedText = await translate(reference[dictName][rKey], {
      from: 'en',
      to: langs[language] as Locale,
    });

    missingTranslations[dictName][rKey] = translatedText;
    missingTranslationsForXlsx.push({
      Label: `${dictName}/${rKey}`,
      English: reference[dictName][rKey],
      [`${language}`]: translatedText,
    });
    if (translatedText !== reference[dictName][rKey]) {
      console.log(`Label #${missingKeys} translated: ${reference[dictName][rKey]} => ${translatedText}`);
    } else {
      console.log(`Couldn't find a translation for ${reference[dictName][rKey]}`);
    }

    missingKeys++;
  }
}

translateMissing().then(() => {
  console.log('Total labels: %d', Object.keys(data[dictName]).length);
  console.log('Missing labels: %d', missingKeys);

  writeFileSync('translation.json', JSON.stringify(data));
  writeFileSync('missing.json', JSON.stringify(missingTranslations));

  if (missingTranslationsForXlsx.length > 0) {
    const wb: reader.WorkBook = reader.utils.book_new();
    reader.utils.book_append_sheet(
      wb,
      reader.utils.json_to_sheet(missingTranslationsForXlsx),
      `${sheetName}-${language}`
    );
    reader.writeFile(wb, 'Missing translations.xlsx');
  }
});
