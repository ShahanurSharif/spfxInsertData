import * as React from 'react';
import type { IInsertDataWebPartProps } from './IInsertDataWebPartProps';
import { Dropdown, IDropdownOption, TextField, PrimaryButton, MessageBar, MessageBarType, IconButton } from '@fluentui/react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { sp } from '@pnp/sp-commonjs';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

const InsertDataWebPart: React.FC<IInsertDataWebPartProps> = (props) => {

  const [Title, setTitle] = React.useState('');
  const [Body, setBody] = React.useState('');
  const [Letter, setLetter] = React.useState('');
  // These are the choices for the dropdown
  const [options, setOptions] = React.useState<IDropdownOption[]>([]);
  // This shows a happy message when you add something
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);
  // These keep track of mistakes in the form
  const [titleError, setTitleError] = React.useState<string | undefined>();
  const [bodyError, setBodyError] = React.useState<string | undefined>();
  const [letterError, setLetterError] = React.useState<string | undefined>();

  const [disabled, setDisabled] = React.useState(true);
  const [showForm, setShowForm] = React.useState(false);

  // FAQ list state
  const [faqItems, setFaqItems] = React.useState<{ Id: number; Title: string; body: string; Letter: string }[]>([]);
  const [editingItem, setEditingItem] = React.useState<{ Id: number; Title: string; body: string; Letter: string } | null>(null);
  const [deletingItem, setDeletingItem] = React.useState<{ Id: number; Title: string } | null>(null);
  const [showDeleteDialog, setShowDeleteDialog] = React.useState(false);
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);

  const validateTitle = (value?: string): boolean => {
    if (!value || value.trim() === '') {
      setTitleError('Title is required');
      return false;
    }
    setTitleError(undefined);
    return true;
  };
  const validateBody = (value?: string): boolean => {
    if (!value || value.trim() === '') {
      setBodyError('Body is required');
      return false;
    }
    setBodyError(undefined);
    return true;
  };
  const validateLetter = (value?: string): boolean => {
    if (!value || value.trim() === '') {
      setLetterError('Letter is required');
      return false;
    }
    setLetterError(undefined);
    return true;
  };

  React.useEffect(() => {
    setDisabled(
      !Title || !Body || !Letter || !!titleError || !!bodyError || !!letterError
    );
  }, [Title, Body, Letter, titleError, bodyError, letterError]);

  // Tell PnPjs how to talk to SharePoint
  React.useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    sp.setup({ spfxContext: props.context as any });
  }, [props.context]);

  // Get the dropdown choices from SharePoint
  React.useEffect(() => {
    const fetchOptions = async (): Promise<void> => {
      try {
        const field = await sp.web.lists
          .getByTitle('FAQ')
          .fields.getByInternalNameOrTitle('Letter')
          .select('Choices').get() as { Choices?: string[] };
        if (field && field.Choices) {
          setOptions(field.Choices.map((choice: string) => ({ key: choice, text: choice })));
        }
      } catch {
        setOptions([
          { key: 'A', text: 'A' },
          { key: 'B', text: 'B' },
          { key: 'C', text: 'C' }
        ]);
      }
    };
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    fetchOptions();
  }, []);

  // Fetch all FAQ items from SharePoint
  const fetchFaqItems = React.useCallback(async () => {
    try {
      const list = sp.web.lists.getByTitle('FAQ');
      const items = await list.items
        .select('Id', 'Title', 'body', 'Letter')
        .orderBy('Id', false)
        .get();
      setFaqItems(items);
    } catch {
      setFaqItems([]);
    }
  }, []);

  React.useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    fetchFaqItems();
  }, [fetchFaqItems, showForm, successMessage]);

  // When you click the button, try to add or update the item
  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    setErrorMessage(null);
    setSuccessMessage(null);
    const isTitleValid = validateTitle(Title);
    const isBodyValid = validateBody(Body);
    const isLetterValid = validateLetter(Letter);
    if (!isTitleValid || !isBodyValid || !isLetterValid) {
      return;
    }
    try {
      if (editingItem) {
        // Update existing item
        await sp.web.lists.getByTitle('FAQ').items.getById(editingItem.Id).update({
          Title,
          body: Body,
          Letter
        });
        setSuccessMessage('Item updated successfully');
      } else {
        // Add new item
        await sp.web.lists.getByTitle('FAQ').items.add({
          Title,
          body: Body,
          Letter
        });
        setSuccessMessage('Item created successfully');
      }
      setTitle('');
      setBody('');
      setLetter('');
      setEditingItem(null);
      setTimeout(() => setSuccessMessage(null), 5000);
      setShowForm(false);
    } catch (error) {
      setErrorMessage('Error creating item');
    }
  };

  // When Edit is clicked, fill the form with the item's values
  const handleEdit = (item: { Id: number; Title: string; body: string; Letter: string }): void => {
    setTitle(item.Title);
    setBody(item.body);
    setLetter(item.Letter);
    setEditingItem(item);
    setShowForm(true);
  };

  // Handle delete icon click
  const handleDelete = (item: { Id: number; Title: string }): void => {
    setDeletingItem(item);
    setShowDeleteDialog(true);
  };

  // Confirm delete
  const confirmDelete = async (): Promise<void> => {
    if (!deletingItem) return;
    setErrorMessage(null);
    setSuccessMessage(null);
    try {
      await sp.web.lists.getByTitle('FAQ').items.getById(deletingItem.Id).delete();
      setSuccessMessage('Item deleted successfully');
      setDeletingItem(null);
      setShowDeleteDialog(false);
    } catch (error) {
      setErrorMessage('Error deleting item');
      setShowDeleteDialog(false);
    }
  };

  // Cancel delete
  const cancelDelete = (): void => {
    setDeletingItem(null);
    setShowDeleteDialog(false);
  };

  // The form you see on the page
  return (
    <div>
      {/* Render success/error message only outside dialog when dialog is closed */}
      {!showForm && successMessage && (
        <div role="alert" data-testid="success-message" style={{ color: 'green', margin: '12px 0' }}>{successMessage}</div>
      )}
      {!showForm && errorMessage && (
        <div role="alert" data-testid="error-message" style={{ color: 'red', margin: '12px 0' }}>{errorMessage}</div>
      )}
      <PrimaryButton text="Create Item" onClick={() => {
        setTitle('');
        setBody('');
        setLetter('');
        setTitleError(undefined);
        setBodyError(undefined);
        setLetterError(undefined);
        setEditingItem(null);
        setShowForm(true);
      }} style={{ marginBottom: 16 }} />
      <Dialog
        hidden={!showForm}
        onDismiss={() => setShowForm(false)}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Create FAQ Item',
        }}
        modalProps={{ isBlocking: false }}
      >
        <form onSubmit={handleSubmit}>
          {/* Render error message only inside dialog when dialog is open */}
          {showForm && errorMessage && (
            <div role="alert" data-testid="error-message" style={{ marginBottom: 8, color: 'red' }}>{errorMessage}</div>
          )}
          <TextField 
            label='Title' 
            id='Title' 
            value={Title} 
            onChange={(event, v) => {
              setTitle(v || '');
              validateTitle(v);
            }} 
            onBlur={() => validateTitle(Title)}
            required
            aria-describedby={titleError ? 'title-error' : undefined}
          />
          {/* Always render error divs for accessibility, but only show text if error exists */}
          <div id="title-error" role="alert" data-testid="title-error" style={{ color: 'red', minHeight: 18 }}>
            {titleError ? titleError : ''}
          </div>
          <TextField 
            label='Body' 
            id='Body' 
            value={Body} 
            onChange={(event, v) => {
              setBody(v || '');
              validateBody(v);
            }} 
            onBlur={() => validateBody(Body)}
            multiline
            required
            aria-describedby={bodyError ? 'body-error' : undefined}
          />
          <div id="body-error" role="alert" data-testid="body-error" style={{ color: 'red', minHeight: 18 }}>
            {bodyError ? bodyError : ''}
          </div>
          <Dropdown 
            label="Letter" 
            id="Letter" 
            options={options} 
            selectedKey={Letter} 
            onChange={(event, option) => {
              setLetter(option ? String(option.key) : '');
              validateLetter(option ? String(option.key) : '');
            }} 
            onBlur={() => validateLetter(Letter)}
            required
            aria-describedby={letterError ? 'letter-error' : undefined}
          />
          <div id="letter-error" role="alert" data-testid="letter-error" style={{ color: 'red', minHeight: 18 }}>
            {letterError ? letterError : ''}
          </div>
          <br />
          <DialogFooter>
            <PrimaryButton text={editingItem ? 'Update' : 'Submit'} type='submit' disabled={disabled} />
            <PrimaryButton text="Cancel" onClick={() => { setShowForm(false); setEditingItem(null); }} />
          </DialogFooter>
        </form>
      </Dialog>
      {/* Delete Confirmation Dialog */}
      <Dialog
        hidden={!showDeleteDialog}
        onDismiss={cancelDelete}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete FAQ Item',
          subText: deletingItem ? `Are you sure you want to delete "${deletingItem.Title}"?` : ''
        }}
        modalProps={{ isBlocking: true }}
      >
        <DialogFooter>
          <PrimaryButton text="Yes, Delete" onClick={confirmDelete} />
          <PrimaryButton text="Cancel" onClick={cancelDelete} />
        </DialogFooter>
      </Dialog>
      {/* FAQ List Table */}
      <h3>FAQ List</h3>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          <tr>
            <th style={{ borderBottom: '1px solid #ccc', textAlign: 'left' }}>Title</th>
            <th style={{ borderBottom: '1px solid #ccc', textAlign: 'left' }}>Body</th>
            <th style={{ borderBottom: '1px solid #ccc', textAlign: 'left' }}>Letter</th>
            <th style={{ borderBottom: '1px solid #ccc', textAlign: 'left' }}>Action</th>
          </tr>
        </thead>
        <tbody>
          {faqItems.map(item => (
            <tr key={item.Id}>
              <td style={{ borderBottom: '1px solid #eee' }}>{item.Title}</td>
              <td style={{ borderBottom: '1px solid #eee' }}>{item.body}</td>
              <td style={{ borderBottom: '1px solid #eee' }}>{item.Letter}</td>
              <td style={{ borderBottom: '1px solid #eee' }}>
                <button
                  type="button"
                  aria-label="Edit"
                  data-testid={`edit-button-${item.Id}`}
                  onClick={() => handleEdit(item)}
                  style={{ marginRight: 8 }}
                >
                  Edit
                </button>
                <button
                  type="button"
                  aria-label="Delete"
                  data-testid={`delete-button-${item.Id}`}
                  onClick={() => handleDelete(item)}
                >
                  Delete
                </button>
              </td>
            </tr>
          ))}
          {faqItems.length === 0 && (
            <tr>
              <td colSpan={4} style={{ textAlign: 'center', color: '#888' }}>No FAQ items found.</td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
};

export default InsertDataWebPart;
