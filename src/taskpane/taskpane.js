/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import '../../assets/icon-16.png';
import '../../assets/icon-32.png';
import '../../assets/icon-80.png';

/* global document, Office, Word */

const form = document.querySelector('.ms-form');

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log(
        'Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.',
      );
    }
    document.getElementById('sideload-msg').style.display = 'none';
    const searchBtn = document.getElementById('search-text');

    searchBtn.addEventListener('click', searchPattern, { once: true });

    form.addEventListener('submit', (e) => {
      e.preventDefault();
      replacePattern();
    });
  }
});

async function searchPattern() {
  return Word.run(async (context) => {
    const result = context.document.body.search('[{][{]*[}][}]', { matchWildcards: true });
    result.load('length');
    result.load('text');
    await context
      .sync()
      .then(() => {
        for (let i = 0; i < result.items.length; i++) {
          form.insertAdjacentHTML(
            'beforeend',
            `<label>${result.items[i].text}<input type="text" /></label>`,
          );
        }
        form.insertAdjacentHTML('beforeend', '<button type="submit">Заменить</button>');
      })
      .then(context.sync);
  });
}

async function replacePattern(ranges) {
  return Word.run(async (context) => {
    const result = context.document.body.search('[{][{]*[}][}]', { matchWildcards: true });
    result.load('text');
    const inputs = document.querySelectorAll('input');
    await context
      .sync()
      .then(() => {
        for (let i = 0; i < result.items.length; i++) {
          console.log(inputs[i].value);
          result.items[i].insertText(inputs[i].value, 'Replace');
        }
      })
      .then(context.sync);
  });
}
