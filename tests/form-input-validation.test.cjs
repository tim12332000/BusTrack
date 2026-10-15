const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const main = fs.readFileSync(path.join(root, '【請複製這個】GoogleAppsScript完整終極版.js'), 'utf8');
const medical = fs.readFileSync(path.join(root, 'apps-script/medical-support.js'), 'utf8');

// Record the rules installed by production form setup. This emulates only the
// Forms validation APIs used here; it does not submit to or change live forms.
function harness() {
  const items = [];
  const form = {
    getItems: () => items.slice(),
    setTitle() { return this; }, setDescription() { return this; },
    moveItem() {},
    deleteItem(item) { items.splice(items.indexOf(item), 1); },
    getDestinationId: () => 'test-sheet',
    isAcceptingResponses: () => true,
    setAcceptingResponses() {},
  };
  for (const [method, type] of [['addTextItem', 'text'], ['addListItem', 'list'], ['addMultipleChoiceItem', 'choice']]) {
    form[method] = () => {
      const item = {
        type, title: '', required: false, validation: null,
        getTitle() { return this.title; }, getType() { return this.type; },
        getId: () => items.indexOf(item), getIndex: () => items.indexOf(item),
        setTitle(value) { this.title = value; return this; },
        setRequired(value) { this.required = value; return this; },
        setValidation(value) { this.validation = value; return this; },
        setChoiceValues() { return this; },
      };
      items.push(item);
      return item;
    };
  }
  const FormApp = {
    ItemType: { TEXT: 'text', LIST: 'list', MULTIPLE_CHOICE: 'choice' },
    openById: () => form,
    createTextValidation() {
      const rules = [];
      return {
        setHelpText() { return this; },
        requireTextMatchesPattern(pattern) {
          rules.push(value => new RegExp(pattern).test(value)); return this;
        },
        requireNumberGreaterThanOrEqualTo(min) {
          rules.push(value => value.trim() !== '' && Number.isFinite(Number(value)) && Number(value) >= min);
          return this;
        },
        build: () => value => rules.every(rule => rule(value)),
      };
    },
  };
  const context = vm.createContext({ FormApp, Logger: { log() {} } });
  vm.runInContext(main, context);
  return { items, form, context };
}

function peopleFields() {
  const h = harness();
  const stop = new Error('Stop after people form setup');
  Object.assign(h.context, {
    SpreadsheetApp: { openById: () => ({}) },
    DriveApp: { getFilesByName: () => ({ hasNext: () => true, next: () => ({ getId: () => 'people' }) }) },
    PropertiesService: { getScriptProperties() { throw stop; } },
  });
  assert.throws(() => h.context.createMultiStationBusSystem(false), error => error === stop);
  return h.items;
}

function medicalFields() {
  const h = harness();
  const stop = new Error('Stop after medical form setup');
  const props = { getProperty: key => key === 'MEDICAL_FORM_ID' ? 'medical' : 'numbered' };
  Object.assign(h.context, {
    PropertiesService: { getScriptProperties: () => props },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => ({}) }) },
  });
  h.form.setDescription = () => { throw stop; };
  vm.runInContext(medical, h.context);
  // The main script's const is authoritative in the combined Apps Script project.
  h.form.getDestinationId = () => vm.runInContext('TARGET_SPREADSHEET_ID', h.context);
  assert.throws(() => h.context.setupMedicalSupport(), error => error === stop);
  return h.items;
}

const people = peopleFields();
const parking = harness();
parking.context.configureParkingRemainingForm(parking.form);
const fields = [
  ['人數', people.find(item => item.title === '4. 人數')],
  ['剩餘汽車位', parking.items.find(item => item.title.includes('剩餘汽車'))],
  ['剩餘機車位', parking.items.find(item => item.title.includes('剩餘機車'))],
  ['衛勤累計人數', medicalFields().find(item => item.title === '累計人數')],
];

function accepts(item, value) {
  assert.ok(item, 'production setup must create the field');
  assert.equal(typeof item.validation, 'function', 'numeric field must have validation');
  if (value.trim() === '') return !item.required;
  return item.validation(value);
}

for (const [name, item] of fields) {
  for (const value of ['0', '1', '25', '1000']) {
    test(`${name} accepts integer ${JSON.stringify(value)}`, () => {
      assert.equal(accepts(item, value), true);
    });
  }
  // Record current form policy; decimal policy is deliberately deferred.
  const decimals = name === '人數' ? [] : ['1.5', '0.5'];
  for (const value of ['abc', '12abc', '.', '..', ...decimals, '-1', '-0.5', '', ' ', '三', '１２', 'NaN', 'Infinity', '=1+1']) {
    test(`${name} rejects invalid input ${JSON.stringify(value)}`, () => {
      assert.equal(accepts(item, value), false, `${name} must reject ${JSON.stringify(value)}`);
    });
  }
}

for (const value of ['1.5', '0.5']) {
  test(`existing people form remains compatible with decimal ${value}`, () => {
    assert.equal(accepts(fields[0][1], value), true);
  });
}

for (const [name, items] of [['人員', people], ['停車場', parking.items]]) {
  test(`${name} notes accept English, symbols and an empty value`, () => {
    const note = items.find(item => item.title.includes('備註'));
    assert.ok(note);
    assert.equal(note.required, false);
    assert.equal(note.validation, null);
  });
}
