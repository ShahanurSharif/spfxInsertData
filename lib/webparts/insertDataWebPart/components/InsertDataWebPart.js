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
import { Dropdown, TextField, PrimaryButton, MessageBar, MessageBarType } from '@fluentui/react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { sp } from '@pnp/sp-commonjs';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
var InsertDataWebPart = function (props) {
    var _a = React.useState(''), Title = _a[0], setTitle = _a[1];
    var _b = React.useState(''), Description = _b[0], setDescription = _b[1];
    var _c = React.useState(''), Letter = _c[0], setLetter = _c[1];
    // These are the choices for the dropdown
    var _d = React.useState([]), options = _d[0], setOptions = _d[1];
    // This shows a happy message when you add something
    var _e = React.useState(null), successMessage = _e[0], setSuccessMessage = _e[1];
    var _f = React.useState(null), errorMessage = _f[0], setErrorMessage = _f[1];
    // These keep track of mistakes in the form
    var _g = React.useState(), titleError = _g[0], setTitleError = _g[1];
    var _h = React.useState(), descriptionError = _h[0], setDescriptionError = _h[1];
    var _j = React.useState(), letterError = _j[0], setLetterError = _j[1];
    var _k = React.useState(true), disabled = _k[0], setDisabled = _k[1];
    var _l = React.useState(false), showForm = _l[0], setShowForm = _l[1];
    // FAQ list state
    var _m = React.useState([]), faqItems = _m[0], setFaqItems = _m[1];
    var _o = React.useState(null), editingItem = _o[0], setEditingItem = _o[1];
    var _p = React.useState(null), deletingItem = _p[0], setDeletingItem = _p[1];
    var _q = React.useState(false), showDeleteDialog = _q[0], setShowDeleteDialog = _q[1];
    var validateTitle = function (value) {
        if (!value || value.trim() === '') {
            setTitleError('Title is required');
            return false;
        }
        setTitleError(undefined);
        return true;
    };
    var validateDescription = function (value) {
        if (!value || value.trim() === '') {
            setDescriptionError('Description is required');
            return false;
        }
        setDescriptionError(undefined);
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
        setDisabled(!Title || !Description || !Letter || !!titleError || !!descriptionError || !!letterError);
    }, [Title, Description, Letter, titleError, descriptionError, letterError]);
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
    // Fetch all FAQ items from SharePoint using 'Body' (capital B)
    var fetchFaqItems = React.useCallback(function () { return __awaiter(void 0, void 0, void 0, function () {
        var list, items, _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _b.trys.push([0, 2, , 3]);
                    list = sp.web.lists.getByTitle('FAQ');
                    return [4 /*yield*/, list.items.select('Id', 'Title', 'Body', 'Letter').orderBy('Id', false).get()];
                case 1:
                    items = _b.sent();
                    setFaqItems(items.map(function (item) { return ({
                        Id: item.Id,
                        Title: item.Title,
                        Description: item.Body, // map 'Body' to Description
                        Letter: item.Letter
                    }); }));
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
        var isTitleValid, isDescriptionValid, isLetterValid, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    e.preventDefault();
                    isTitleValid = validateTitle(Title);
                    isDescriptionValid = validateDescription(Description);
                    isLetterValid = validateLetter(Letter);
                    if (!isTitleValid || !isDescriptionValid || !isLetterValid) {
                        return [2 /*return*/];
                    }
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, , 7]);
                    if (!editingItem) return [3 /*break*/, 3];
                    return [4 /*yield*/, sp.web.lists.getByTitle('FAQ').items.getById(editingItem.Id).update({
                            Title: Title,
                            Body: Description,
                            Letter: Letter
                        })];
                case 2:
                    _a.sent();
                    setSuccessMessage('Item updated successfully');
                    setErrorMessage(null);
                    setTitle('');
                    setDescription('');
                    setLetter('');
                    setEditingItem(null);
                    setTimeout(function () {
                        setSuccessMessage(null);
                        setShowForm(false); // Close dialog after update
                    }, 3000);
                    return [3 /*break*/, 5];
                case 3: return [4 /*yield*/, sp.web.lists.getByTitle('FAQ').items.add({
                        Title: Title,
                        Body: Description,
                        Letter: Letter
                    })];
                case 4:
                    _a.sent();
                    setSuccessMessage('Item created successfully');
                    setErrorMessage(null);
                    setTitle('');
                    setDescription('');
                    setLetter('');
                    setTimeout(function () {
                        setSuccessMessage(null);
                    }, 3000);
                    _a.label = 5;
                case 5: return [3 /*break*/, 7];
                case 6:
                    error_1 = _a.sent();
                    setErrorMessage('Error creating item');
                    setSuccessMessage(null);
                    return [3 /*break*/, 7];
                case 7: return [2 /*return*/];
            }
        });
    }); };
    // When Edit is clicked, fill the form with the item's values
    var handleEdit = function (item) {
        setTitle(item.Title);
        setDescription(item.Description); // Description is already mapped from 'body'
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
        var _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    if (!deletingItem)
                        return [2 /*return*/];
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, sp.web.lists.getByTitle('FAQ').items.getById(deletingItem.Id).delete()];
                case 2:
                    _b.sent();
                    setSuccessMessage('Item deleted successfully');
                    setDeletingItem(null);
                    setShowDeleteDialog(false);
                    return [3 /*break*/, 4];
                case 3:
                    _a = _b.sent();
                    setErrorMessage('Error deleting item');
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
    // Success message auto-dismiss effect (only for create, update handled in handleSubmit)
    React.useEffect(function () {
        if (!showForm && successMessage && !editingItem) {
            var timer_1 = setTimeout(function () { return setSuccessMessage(null); }, 3000);
            return function () { return clearTimeout(timer_1); };
        }
    }, [successMessage, showForm, editingItem]);
    // The form you see on the page
    return (React.createElement("div", null,
        React.createElement(PrimaryButton, { text: "Create Item", onClick: function () {
                setTitle('');
                setDescription('');
                setLetter('');
                setTitleError(undefined);
                setDescriptionError(undefined);
                setLetterError(undefined);
                setEditingItem(null);
                setShowForm(true);
            }, style: { marginBottom: 16 } }),
        !showForm && successMessage && (React.createElement(MessageBar, { messageBarType: MessageBarType.success, isMultiline: false, "data-testid": "success-message", role: "alert", onDismiss: undefined, styles: { root: { position: 'absolute', top: 24, left: '50%', transform: 'translateX(-50%)', zIndex: 9999, minWidth: 320, maxWidth: 480, textAlign: 'center' } } }, successMessage)),
        !showForm && errorMessage && (React.createElement(MessageBar, { messageBarType: MessageBarType.error, isMultiline: false, "data-testid": "error-message", role: "alert", onDismiss: function () { return setErrorMessage(null); }, dismissButtonAriaLabel: "Dismiss error message", styles: { root: { position: 'absolute', top: 24, left: '50%', transform: 'translateX(-50%)', zIndex: 9999, minWidth: 320, maxWidth: 480, textAlign: 'center' } } }, errorMessage)),
        React.createElement(Dialog, { hidden: !showForm, onDismiss: function () { return setShowForm(false); }, dialogContentProps: {
                type: DialogType.largeHeader,
                title: 'Create FAQ Item',
            }, modalProps: { isBlocking: false } },
            React.createElement("form", { onSubmit: handleSubmit },
                showForm && errorMessage && (React.createElement("div", { role: "alert", "data-testid": "error-message", style: { marginBottom: 8, color: 'red' } }, errorMessage)),
                React.createElement(TextField, { label: 'Title', id: 'Title', value: Title, onChange: function (event, v) {
                        setTitle(v || '');
                        validateTitle(v);
                    }, onBlur: function () { return validateTitle(Title); }, required: true, "aria-describedby": titleError ? 'title-error' : undefined }),
                React.createElement("div", { id: "title-error", role: "alert", "data-testid": "title-error", style: { color: 'red', minHeight: 18 } }, titleError ? titleError : ''),
                React.createElement(TextField, { label: 'Description', id: 'Description', value: Description, onChange: function (event, v) {
                        setDescription(v || '');
                        validateDescription(v);
                    }, onBlur: function () { return validateDescription(Description); }, multiline: true, required: true, "aria-describedby": descriptionError ? 'description-error' : undefined }),
                React.createElement("div", { id: "description-error", role: "alert", "data-testid": "description-error", style: { color: 'red', minHeight: 18 } }, descriptionError ? descriptionError : ''),
                React.createElement(Dropdown, { label: "Letter", id: "Letter", options: options, selectedKey: Letter, onChange: function (event, option) {
                        setLetter(option ? String(option.key) : '');
                        validateLetter(option ? String(option.key) : '');
                    }, onBlur: function () { return validateLetter(Letter); }, required: true, "aria-describedby": letterError ? 'letter-error' : undefined }),
                React.createElement("div", { id: "letter-error", role: "alert", "data-testid": "letter-error", style: { color: 'red', minHeight: 18 } }, letterError ? letterError : ''),
                React.createElement("br", null),
                React.createElement(DialogFooter, null,
                    React.createElement(PrimaryButton, { text: editingItem ? 'Update' : 'Submit', type: 'submit', disabled: disabled }),
                    React.createElement(PrimaryButton, { text: "Close", onClick: function () { setShowForm(false); setEditingItem(null); } })))),
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
                    React.createElement("th", { style: { borderBottom: '1px solid #ccc', textAlign: 'left' } }, "Description"),
                    React.createElement("th", { style: { borderBottom: '1px solid #ccc', textAlign: 'left' } }, "Letter"),
                    React.createElement("th", { style: { borderBottom: '1px solid #ccc', textAlign: 'left' } }, "Action"))),
            React.createElement("tbody", null,
                faqItems.map(function (item) { return (React.createElement("tr", { key: item.Id },
                    React.createElement("td", { style: { borderBottom: '1px solid #eee' } }, item.Title),
                    React.createElement("td", { style: { borderBottom: '1px solid #eee' } }, item.Description),
                    React.createElement("td", { style: { borderBottom: '1px solid #eee' } }, item.Letter),
                    React.createElement("td", { style: { borderBottom: '1px solid #eee' } },
                        React.createElement("button", { type: "button", "aria-label": "Edit", "data-testid": "edit-button-".concat(item.Id), onClick: function () { return handleEdit(item); }, style: { marginRight: 8 } }, "Edit"),
                        React.createElement("button", { type: "button", "aria-label": "Delete", "data-testid": "delete-button-".concat(item.Id), onClick: function () { return handleDelete({ Id: item.Id, Title: item.Title }); } }, "Delete")))); }),
                faqItems.length === 0 && (React.createElement("tr", null,
                    React.createElement("td", { colSpan: 4, style: { textAlign: 'center', color: '#888' } }, "No FAQ items found.")))))));
};
export default InsertDataWebPart;
//# sourceMappingURL=InsertDataWebPart.js.map