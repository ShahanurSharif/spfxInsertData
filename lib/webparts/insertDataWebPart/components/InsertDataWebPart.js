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
import * as React from 'react';
import { Dropdown, TextField, PrimaryButton, MessageBar, MessageBarType, IconButton } from '@fluentui/react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { sp } from '@pnp/sp-commonjs';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
var InsertDataWebPart = function (props) {
    var _a = React.useState(''), Title = _a[0], setTitle = _a[1];
    var _b = React.useState(''), Body = _b[0], setBody = _b[1];
    var _c = React.useState(''), Letter = _c[0], setLetter = _c[1];
    // These are the choices for the dropdown
    var _d = React.useState([]), options = _d[0], setOptions = _d[1];
    // This shows a happy message when you add something
    var _e = React.useState(null), successMessage = _e[0], setSuccessMessage = _e[1];
    // These keep track of mistakes in the form
    var _f = React.useState(), titleError = _f[0], setTitleError = _f[1];
    var _g = React.useState(), bodyError = _g[0], setBodyError = _g[1];
    var _h = React.useState(), letterError = _h[0], setLetterError = _h[1];
    var _j = React.useState(true), disabled = _j[0], setDisabled = _j[1];
    var _k = React.useState(false), showForm = _k[0], setShowForm = _k[1];
    // FAQ list state
    var _l = React.useState([]), faqItems = _l[0], setFaqItems = _l[1];
    var _m = React.useState(null), editingItem = _m[0], setEditingItem = _m[1];
    var _o = React.useState(null), deletingItem = _o[0], setDeletingItem = _o[1];
    var _p = React.useState(false), showDeleteDialog = _p[0], setShowDeleteDialog = _p[1];
    var validateTitle = function (value) {
        if (!value || value.trim() === '') {
            setTitleError('Title is required');
            return false;
        }
        setTitleError(undefined);
        return true;
    };
    var validateBody = function (value) {
        if (!value || value.trim() === '') {
            setBodyError('Body is required');
            return false;
        }
        setBodyError(undefined);
        return true;
    };
    var validateLetter = function (value) {
        if (!value || value.trim() === '') {
            setLetterError('Letter is required');
            return false;
        }
        setLetterError(undefined);
        return true;
    };
    React.useEffect(function () {
        setDisabled(!Title || !Body || !Letter || !!titleError || !!bodyError || !!letterError);
    }, [Title, Body, Letter, titleError, bodyError, letterError]);
    // Tell PnPjs how to talk to SharePoint
    React.useEffect(function () {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        sp.setup({ spfxContext: props.context });
    }, [props.context]);
    // Get the dropdown choices from SharePoint
    React.useEffect(function () {
        var fetchOptions = function () { return __awaiter(void 0, void 0, void 0, function () {
            var field, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle('FAQ')
                                .fields.getByInternalNameOrTitle('Letter')
                                .select('Choices').get()];
                    case 1:
                        field = _b.sent();
                        if (field && field.Choices) {
                            setOptions(field.Choices.map(function (choice) { return ({ key: choice, text: choice }); }));
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        _a = _b.sent();
                        setOptions([
                            { key: 'A', text: 'A' },
                            { key: 'B', text: 'B' },
                            { key: 'C', text: 'C' }
                        ]);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        fetchOptions();
    }, []);
    // Fetch all FAQ items from SharePoint
    var fetchFaqItems = React.useCallback(function () { return __awaiter(void 0, void 0, void 0, function () {
        var list, items, _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _b.trys.push([0, 2, , 3]);
                    list = sp.web.lists.getByTitle('FAQ');
                    return [4 /*yield*/, list.items
                            .select('Id', 'Title', 'body', 'Letter')
                            .orderBy('Id', false)
                            .get()];
                case 1:
                    items = _b.sent();
                    setFaqItems(items);
                    return [3 /*break*/, 3];
                case 2:
                    _a = _b.sent();
                    setFaqItems([]);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    }); }, []);
    React.useEffect(function () {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        fetchFaqItems();
    }, [fetchFaqItems, showForm, successMessage]);
    // When you click the button, try to add or update the item
    var handleSubmit = function (e) { return __awaiter(void 0, void 0, void 0, function () {
        var isTitleValid, isBodyValid, isLetterValid, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    e.preventDefault();
                    isTitleValid = validateTitle(Title);
                    isBodyValid = validateBody(Body);
                    isLetterValid = validateLetter(Letter);
                    if (!isTitleValid || !isBodyValid || !isLetterValid) {
                        return [2 /*return*/];
                    }
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, , 7]);
                    if (!editingItem) return [3 /*break*/, 3];
                    // Update existing item
                    return [4 /*yield*/, sp.web.lists.getByTitle('FAQ').items.getById(editingItem.Id).update({
                            Title: Title,
                            body: Body,
                            Letter: Letter
                        })];
                case 2:
                    // Update existing item
                    _a.sent();
                    setSuccessMessage('FAQ item updated successfully!');
                    return [3 /*break*/, 5];
                case 3: 
                // Add new item
                return [4 /*yield*/, sp.web.lists.getByTitle('FAQ').items.add({
                        Title: Title,
                        body: Body,
                        Letter: Letter
                    })];
                case 4:
                    // Add new item
                    _a.sent();
                    setSuccessMessage('FAQ item added successfully!');
                    _a.label = 5;
                case 5:
                    setTitle('');
                    setBody('');
                    setLetter('');
                    setEditingItem(null);
                    setTimeout(function () { return setSuccessMessage(null); }, 5000);
                    setShowForm(false);
                    return [3 /*break*/, 7];
                case 6:
                    error_1 = _a.sent();
                    alert('Error saving FAQ item: ' + error_1);
                    return [3 /*break*/, 7];
                case 7: return [2 /*return*/];
            }
        });
    }); };
    // When Edit is clicked, fill the form with the item's values
    var handleEdit = function (item) {
        setTitle(item.Title);
        setBody(item.body);
        setLetter(item.Letter);
        setEditingItem(item);
        setShowForm(true);
    };
    // Handle delete icon click
    var handleDelete = function (item) {
        setDeletingItem(item);
        setShowDeleteDialog(true);
    };
    // Confirm delete
    var confirmDelete = function () { return __awaiter(void 0, void 0, void 0, function () {
        var error_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!deletingItem)
                        return [2 /*return*/];
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, sp.web.lists.getByTitle('FAQ').items.getById(deletingItem.Id).delete()];
                case 2:
                    _a.sent();
                    setSuccessMessage('FAQ item deleted successfully!');
                    setDeletingItem(null);
                    setShowDeleteDialog(false);
                    return [3 /*break*/, 4];
                case 3:
                    error_2 = _a.sent();
                    alert('Error deleting FAQ item: ' + error_2);
                    setShowDeleteDialog(false);
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    // Cancel delete
    var cancelDelete = function () {
        setDeletingItem(null);
        setShowDeleteDialog(false);
    };
    // The form you see on the page
    return (React.createElement("div", null,
        React.createElement(PrimaryButton, { text: "Create Item", onClick: function () {
                setTitle('');
                setBody('');
                setLetter('');
                setTitleError(undefined);
                setBodyError(undefined);
                setLetterError(undefined);
                setEditingItem(null);
                setShowForm(true);
            }, style: { marginBottom: 16 } }),
        React.createElement(Dialog, { hidden: !showForm, onDismiss: function () { return setShowForm(false); }, dialogContentProps: {
                type: DialogType.largeHeader,
                title: 'Create FAQ Item',
            }, modalProps: { isBlocking: false } },
            React.createElement("form", { onSubmit: handleSubmit },
                successMessage && (React.createElement(MessageBar, { messageBarType: MessageBarType.success, isMultiline: false, onDismiss: function () { return setSuccessMessage(null); } }, successMessage)),
                React.createElement(TextField, { label: 'Title', id: 'Title', value: Title, onChange: function (event, v) {
                        setTitle(v || '');
                        validateTitle(v);
                    }, onBlur: function () { return validateTitle(Title); }, errorMessage: titleError, required: true }),
                React.createElement(TextField, { label: 'Body', id: 'Body', value: Body, onChange: function (event, v) {
                        setBody(v || '');
                        validateBody(v);
                    }, onBlur: function () { return validateBody(Body); }, errorMessage: bodyError, multiline: true, required: true }),
                React.createElement(Dropdown, { label: "Letter", id: "Letter", options: options, selectedKey: Letter, onChange: function (event, option) {
                        setLetter(option ? String(option.key) : '');
                        validateLetter(option ? String(option.key) : '');
                    }, onBlur: function () { return validateLetter(Letter); }, errorMessage: letterError, required: true }),
                React.createElement("br", null),
                React.createElement(DialogFooter, null,
                    React.createElement(PrimaryButton, { text: editingItem ? 'Update' : 'Submit', type: 'submit', disabled: disabled }),
                    React.createElement(PrimaryButton, { text: "Cancel", onClick: function () { setShowForm(false); setEditingItem(null); } })))),
        React.createElement(Dialog, { hidden: !showDeleteDialog, onDismiss: cancelDelete, dialogContentProps: {
                type: DialogType.normal,
                title: 'Delete FAQ Item',
                subText: deletingItem ? "Are you sure you want to delete \"".concat(deletingItem.Title, "\"?") : ''
            }, modalProps: { isBlocking: true } },
            React.createElement(DialogFooter, null,
                React.createElement(PrimaryButton, { text: "Yes, Delete", onClick: confirmDelete }),
                React.createElement(PrimaryButton, { text: "Cancel", onClick: cancelDelete }))),
        React.createElement("h3", null, "FAQ List"),
        React.createElement("table", { style: { width: '100%', borderCollapse: 'collapse' } },
            React.createElement("thead", null,
                React.createElement("tr", null,
                    React.createElement("th", { style: { borderBottom: '1px solid #ccc', textAlign: 'left' } }, "Title"),
                    React.createElement("th", { style: { borderBottom: '1px solid #ccc', textAlign: 'left' } }, "Body"),
                    React.createElement("th", { style: { borderBottom: '1px solid #ccc', textAlign: 'left' } }, "Letter"),
                    React.createElement("th", { style: { borderBottom: '1px solid #ccc', textAlign: 'left' } }, "Action"))),
            React.createElement("tbody", null,
                faqItems.map(function (item) { return (React.createElement("tr", { key: item.Id },
                    React.createElement("td", { style: { borderBottom: '1px solid #eee' } }, item.Title),
                    React.createElement("td", { style: { borderBottom: '1px solid #eee' } }, item.body),
                    React.createElement("td", { style: { borderBottom: '1px solid #eee' } }, item.Letter),
                    React.createElement("td", { style: { borderBottom: '1px solid #eee' } },
                        React.createElement(IconButton, { iconProps: { iconName: 'Edit', style: { color: 'green' } }, title: "Edit", ariaLabel: "Edit", onClick: function () { return handleEdit(item); } }),
                        React.createElement(IconButton, { iconProps: { iconName: 'Delete', style: { color: 'red' } }, title: "Delete", ariaLabel: "Delete", onClick: function () { return handleDelete(item); } })))); }),
                faqItems.length === 0 && (React.createElement("tr", null,
                    React.createElement("td", { colSpan: 3, style: { textAlign: 'center', color: '#888' } }, "No FAQ items found.")))))));
};
export default InsertDataWebPart;
//# sourceMappingURL=InsertDataWebPart.js.map