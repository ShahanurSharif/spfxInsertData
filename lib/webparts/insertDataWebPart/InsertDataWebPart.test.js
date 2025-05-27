var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
/*
for single test use: npx jest src/webparts/insertDataWebPart/InsertDataWebPart.test.tsx -t 'opens the form dialog when Create Item is clicked'
*/
/// <reference types="jest" />
import { render, screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import '@testing-library/jest-dom';
import InsertDataWebPart from './components/InsertDataWebPart';
import { faker } from '@faker-js/faker';
import * as React from 'react';
// function selectDropdownOption(arg0: string, arg1: string) {
//   throw new Error('Function not implemented.');
// }
var mockContext = {};
var mockProps = {
    context: mockContext,
    description: '',
    isDarkTheme: false,
    environmentMessage: '',
    hasTeamsContext: false,
    userDisplayName: ''
};
describe('InsertDataWebPart', function () {
    it('renders Create Item button', function () {
        render(React.createElement(InsertDataWebPart, __assign({}, mockProps)));
        expect(screen.getByText('Create Item')).toBeInTheDocument();
    });
    it('opens the form dialog when Create Item is clicked', function () {
        render(React.createElement(InsertDataWebPart, __assign({}, mockProps)));
        userEvent.click(screen.getByText('Create Item'));
        expect(screen.getByText('Create FAQ Item')).toBeInTheDocument();
    });
    it('shows validation errors if required fields are empty after blur', function () { return __awaiter(void 0, void 0, void 0, function () {
        var titleInput, bodyInput, letterDropdown, alerts;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    render(React.createElement(InsertDataWebPart, __assign({}, mockProps)));
                    userEvent.click(screen.getByText('Create Item'));
                    titleInput = screen.getByLabelText('Title');
                    titleInput.focus();
                    titleInput.blur();
                    bodyInput = screen.getByLabelText('Body');
                    bodyInput.focus();
                    bodyInput.blur();
                    letterDropdown = screen.getByLabelText('Letter');
                    letterDropdown.focus();
                    letterDropdown.blur();
                    return [4 /*yield*/, screen.findAllByRole('alert')];
                case 1:
                    alerts = _a.sent();
                    expect(alerts.some(function (a) { var _a; return (_a = a.textContent) === null || _a === void 0 ? void 0 : _a.match(/Title is required/i); })).toBe(true);
                    expect(alerts.some(function (a) { var _a; return (_a = a.textContent) === null || _a === void 0 ? void 0 : _a.match(/Body is required/i); })).toBe(true);
                    expect(alerts.some(function (a) { var _a; return (_a = a.textContent) === null || _a === void 0 ? void 0 : _a.match(/Letter is required/i); })).toBe(true);
                    return [2 /*return*/];
            }
        });
    }); });
    function selectDropdownOption(label, optionText) {
        return __awaiter(this, void 0, void 0, function () {
            var listbox, option;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        userEvent.click(screen.getByLabelText(label));
                        return [4 /*yield*/, screen.findByRole('listbox')];
                    case 1:
                        listbox = _a.sent();
                        option = within(listbox).getByText(optionText);
                        userEvent.click(option);
                        return [2 /*return*/];
                }
            });
        });
    }
    it('create an item successfully', function () { return __awaiter(void 0, void 0, void 0, function () {
        var fakeTitle;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    render(React.createElement(InsertDataWebPart, __assign({}, mockProps)));
                    userEvent.click(screen.getByText('Create Item'));
                    fakeTitle = faker.lorem.sentence();
                    userEvent.type(screen.getByLabelText('Title'), fakeTitle);
                    userEvent.type(screen.getByLabelText('Body'), fakeTitle);
                    return [4 /*yield*/, selectDropdownOption('Letter', 'A')];
                case 1:
                    _a.sent();
                    userEvent.click(screen.getByText('Submit'));
                    // Wait for the MessageBar to appear, then check for the success message
                    // const messageBar = await screen.findByTestId('success-message', {}, { timeout: 5000 });
                    // expect(messageBar).toBeInTheDocument();
                    // expect(within(messageBar).getByText(/item created successfully/i)).toBeInTheDocument();
                    expect(screen.getByText(fakeTitle)).toBeInTheDocument();
                    expect(screen.getByText('A')).toBeInTheDocument();
                    return [2 /*return*/];
            }
        });
    }); });
    it('shows error message if item creation fails', function () { return __awaiter(void 0, void 0, void 0, function () {
        var _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    render(React.createElement(InsertDataWebPart, __assign({}, mockProps)));
                    userEvent.click(screen.getByText('Create Item'));
                    userEvent.type(screen.getByLabelText('Title'), 'Test Title');
                    userEvent.type(screen.getByLabelText('Body'), 'Test Body');
                    return [4 /*yield*/, selectDropdownOption('Letter', 'A')];
                case 1:
                    _b.sent();
                    userEvent.click(screen.getByText('Submit'));
                    _a = expect;
                    return [4 /*yield*/, screen.findByText('Error creating item')];
                case 2:
                    _a.apply(void 0, [_b.sent()]).toBeInTheDocument();
                    return [2 /*return*/];
            }
        });
    }); });
    it('edits an item successfully', function () { return __awaiter(void 0, void 0, void 0, function () {
        var fakeTitle, itemRow;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    // Reset fetch mock from previous tests to ensure success
                    window.fetch = jest.fn(function () {
                        return Promise.resolve({
                            ok: true,
                            json: function () { return Promise.resolve({ Id: 1, Title: 'Test Title', Body: 'Test Body', Letter: 'A' }); }
                        });
                    });
                    render(React.createElement(InsertDataWebPart, __assign({}, mockProps)));
                    userEvent.click(screen.getByText('Create Item'));
                    fakeTitle = faker.lorem.sentence();
                    userEvent.type(screen.getByLabelText('Title'), fakeTitle);
                    userEvent.type(screen.getByLabelText('Body'), fakeTitle);
                    return [4 /*yield*/, selectDropdownOption('Letter', 'A')];
                case 1:
                    _a.sent();
                    userEvent.click(screen.getByText('Submit'));
                    userEvent.click(screen.getByText('Close'));
                    // Wait for the form dialog to close and the item to appear in the list
                    return [4 /*yield*/, screen.findByText(fakeTitle)];
                case 2:
                    // Wait for the form dialog to close and the item to appear in the list
                    _a.sent();
                    return [4 /*yield*/, screen.findByText(fakeTitle)];
                case 3:
                    itemRow = _a.sent();
                    expect(itemRow).toBeInTheDocument();
                    return [2 /*return*/];
            }
        });
    }); });
    it('deletes an item successfully', function () { return __awaiter(void 0, void 0, void 0, function () {
        var fakeTitle;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    render(React.createElement(InsertDataWebPart, __assign({}, mockProps)));
                    userEvent.click(screen.getByText('Create Item'));
                    fakeTitle = faker.lorem.sentence();
                    userEvent.type(screen.getByLabelText('Title'), fakeTitle);
                    userEvent.type(screen.getByLabelText('Body'), faker.lorem.paragraph());
                    return [4 /*yield*/, selectDropdownOption('Letter', 'A')];
                case 1:
                    _a.sent();
                    userEvent.click(screen.getByText('Submit'));
                    // Wait for the dialog to close and the item to appear in the list
                    return [4 /*yield*/, screen.findByText(fakeTitle)];
                case 2:
                    // Wait for the dialog to close and the item to appear in the list
                    _a.sent();
                    // Now delete the item
                    userEvent.click(screen.getByText('Delete'));
                    expect(screen.queryByText(fakeTitle)).not.toBeInTheDocument();
                    return [2 /*return*/];
            }
        });
    }); });
    it('shows error message if item deletion fails', function () { return __awaiter(void 0, void 0, void 0, function () {
        var _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    window.fetch = jest.fn(function () { return Promise.reject(new Error('Network error')); });
                    render(React.createElement(InsertDataWebPart, __assign({}, mockProps)));
                    userEvent.click(screen.getByText('Create Item'));
                    userEvent.type(screen.getByLabelText('Title'), 'Test Title');
                    userEvent.type(screen.getByLabelText('Body'), 'Test Body');
                    return [4 /*yield*/, selectDropdownOption('Letter', 'A')];
                case 1:
                    _b.sent();
                    userEvent.click(screen.getByText('Submit'));
                    // Now try to delete the item
                    userEvent.click(screen.getByText('Delete'));
                    _a = expect;
                    return [4 /*yield*/, screen.findByText('Error deleting item')];
                case 2:
                    _a.apply(void 0, [_b.sent()]).toBeInTheDocument();
                    return [2 /*return*/];
            }
        });
    }); });
    it('renders edit and delete buttons with correct data-testid and allows extracting item Id', function () { return __awaiter(void 0, void 0, void 0, function () {
        var fakeTitle, editButton, dataTestId, idMatch, itemId, deleteButton;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    render(React.createElement(InsertDataWebPart, __assign({}, mockProps)));
                    userEvent.click(screen.getByText('Create Item'));
                    fakeTitle = faker.lorem.sentence();
                    userEvent.type(screen.getByLabelText('Title'), fakeTitle);
                    userEvent.type(screen.getByLabelText('Body'), 'Test body');
                    return [4 /*yield*/, selectDropdownOption('Letter', 'A')];
                case 1:
                    _a.sent();
                    userEvent.click(screen.getByText('Submit'));
                    return [4 /*yield*/, screen.findByTestId(/^edit-button-\d+$/)];
                case 2:
                    editButton = _a.sent();
                    expect(editButton).toBeInTheDocument();
                    dataTestId = editButton.getAttribute('data-testid');
                    expect(dataTestId).toMatch(/^edit-button-\d+$/);
                    idMatch = dataTestId === null || dataTestId === void 0 ? void 0 : dataTestId.match(/^edit-button-(\d+)$/);
                    expect(idMatch).not.toBeNull();
                    itemId = idMatch ? Number(idMatch[1]) : null;
                    expect(typeof itemId).toBe('number');
                    expect(itemId).toBeGreaterThan(0);
                    deleteButton = screen.getByTestId("delete-button-".concat(itemId));
                    expect(deleteButton).toBeInTheDocument();
                    return [2 /*return*/];
            }
        });
    }); });
});
//# sourceMappingURL=InsertDataWebPart.test.js.map