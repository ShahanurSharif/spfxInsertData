import * as React from 'react';
import type { IInsertDataWebPartProps } from './IInsertDataWebPartProps';
import { Dropdown, IDropdownOption, TextField, PrimaryButton, MessageBar, MessageBarType } from '@fluentui/react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { sp } from '@pnp/sp-commonjs';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

const InsertDataWebPart: React.FC<IInsertDataWebPartProps> = (props) => {

  const [Title, setTitle] = React.useState('');
  const [Description, setDescription] = React.useState('');
  const [Letter, setLetter] = React.useState('');
  // These are the choices for the dropdown
  const [options, setOptions] = React.useState<IDropdownOption[]>([]);
  // This shows a happy message when you add something
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
  // These keep track of mistakes in the form
  const [titleError, setTitleError] = React.useState<string | undefined>();
  const [descriptionError, setDescriptionError] = React.useState<string | undefined>();
  const [letterError, setLetterError] = React.useState<string | undefined>();

  const [disabled, setDisabled] = React.useState(true);
  const [showForm, setShowForm] = React.useState(false);

  // FAQ list state
  const [faqItems, setFaqItems] = React.useState<{ Id: number; Title: string; Description: string; Letter: string }[]>([]);
  const [editingItem, setEditingItem] = React.useState<{ Id: number; Title: string; Description: string; Letter: string } | null>(null);
  const [deletingItem, setDeletingItem] = React.useState<{ Id: number; Title: string } | null>(null);
  const [showDeleteDialog, setShowDeleteDialog] = React.useState(false);

  const validateTitle = (value?: string): boolean => {
    if (!value || value.trim() === '') {
      setTitleError('Title is required');
      return false;
    }
    setTitleError(undefined);
    return true;
  };
  const validateDescription = (value?: string): boolean => {
    if (!value || value.trim() === '') {
      setDescriptionError('Description is required');
      return false;
    }
    setDescriptionError(undefined);
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
      !Title || !Description || !Letter || !!titleError || !!descriptionError || !!letterError
    );
  }, [Title, Description, Letter, titleError, descriptionError, letterError]);

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

  // Fetch all FAQ items from SharePoint using 'Body' (capital B)
  const fetchFaqItems = React.useCallback(async () => {
    try {
      const list = sp.web.lists.getByTitle('FAQ');
      const items = await list.items.select('Id', 'Title', 'Body', 'Letter').orderBy('Id', false).get();
      setFaqItems(items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Description: item.Body, // map 'Body' to Description
        Letter: item.Letter
      })));
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
    const isTitleValid = validateTitle(Title);
    const isDescriptionValid = validateDescription(Description);
    const isLetterValid = validateLetter(Letter);
    if (!isTitleValid || !isDescriptionValid || !isLetterValid) {
      return;
    }
    try {
      if (editingItem) {
        await sp.web.lists.getByTitle('FAQ').items.getById(editingItem.Id).update({
          Title,
          Body: Description,
          Letter
        });
        setSuccessMessage('Item updated successfully');
        setErrorMessage(null);
        setTitle('');
        setDescription('');
        setLetter('');
        setEditingItem(null);
        setTimeout(() => {
          setSuccessMessage(null);
          setShowForm(false); // Close dialog after update
        }, 3000);
      } else {
        await sp.web.lists.getByTitle('FAQ').items.add({
          Title,
          Body: Description,
          Letter
        });
        setSuccessMessage('Item created successfully');
        setErrorMessage(null);
        setTitle('');
        setDescription('');
        setLetter('');
        setTimeout(() => {
          setSuccessMessage(null);
        }, 3000);
      }
    } catch (error) {
      setErrorMessage('Error creating item');
      setSuccessMessage(null);
    }
  };

  // When Edit is clicked, fill the form with the item's values
  const handleEdit = (item: { Id: number; Title: string; Description: string; Letter: string }): void => {
    setTitle(item.Title);
    setDescription(item.Description); // Description is already mapped from 'body'
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
    try {
      await sp.web.lists.getByTitle('FAQ').items.getById(deletingItem.Id).delete();
      setSuccessMessage('Item deleted successfully');
      setDeletingItem(null);
      setShowDeleteDialog(false);
    } catch {
      setErrorMessage('Error deleting item');
      setShowDeleteDialog(false);
    }
  };

  // Cancel delete
  const cancelDelete = (): void => {
    setDeletingItem(null);
    setShowDeleteDialog(false);
  };

  // Success message auto-dismiss effect (only for create, update handled in handleSubmit)
  React.useEffect(() => {
    if (!showForm && successMessage && !editingItem) {
      const timer = setTimeout(() => setSuccessMessage(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [successMessage, showForm, editingItem]);

  // The form you see on the page
  return (
    <div>
      <PrimaryButton text="Create Item" onClick={() => {
        setTitle('');
        setDescription('');
        setLetter('');
        setTitleError(undefined);
        setDescriptionError(undefined);
        setLetterError(undefined);
        setEditingItem(null);
        setShowForm(true);
      }} style={{ marginBottom: 16 }} />
      {/* Render success message outside dialog, auto-dismiss after 3s, styled absolute top center */}
      {!showForm && successMessage && (
        <MessageBar
          messageBarType={MessageBarType.success}
          isMultiline={false}
          data-testid="success-message"
          role="alert"
          onDismiss={undefined}
          styles={{ root: { position: 'absolute', top: 24, left: '50%', transform: 'translateX(-50%)', zIndex: 9999, minWidth: 320, maxWidth: 480, textAlign: 'center' } }}
        >
          {successMessage}
        </MessageBar>
      )}
      {/* Render error message outside dialog, dismissible by user, styled absolute top center */}
      {!showForm && errorMessage && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          data-testid="error-message"
          role="alert"
          onDismiss={() => setErrorMessage(null)}
          dismissButtonAriaLabel="Dismiss error message"
          styles={{ root: { position: 'absolute', top: 24, left: '50%', transform: 'translateX(-50%)', zIndex: 9999, minWidth: 320, maxWidth: 480, textAlign: 'center' } }}
        >
          {errorMessage}
        </MessageBar>
      )}
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
            label='Description' 
            id='Description' 
            value={Description} 
            onChange={(event, v) => {
              setDescription(v || '');
              validateDescription(v);
            }} 
            onBlur={() => validateDescription(Description)}
            multiline
            required
            aria-describedby={descriptionError ? 'description-error' : undefined}
          />
          <div id="description-error" role="alert" data-testid="description-error" style={{ color: 'red', minHeight: 18 }}>
            {descriptionError ? descriptionError : ''}
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
          {/* Only render error if present */}
          <div id="letter-error" role="alert" data-testid="letter-error" style={{ color: 'red', minHeight: 18 }}>
            {letterError ? letterError : ''}
          </div>
          <br />
          <DialogFooter>
            <PrimaryButton text={editingItem ? 'Update' : 'Submit'} type='submit' disabled={disabled} />
            <PrimaryButton text="Close" onClick={() => { setShowForm(false); setEditingItem(null); }} />
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
            <th style={{ borderBottom: '1px solid #ccc', textAlign: 'left' }}>Description</th>
            <th style={{ borderBottom: '1px solid #ccc', textAlign: 'left' }}>Letter</th>
            <th style={{ borderBottom: '1px solid #ccc', textAlign: 'left' }}>Action</th>
          </tr>
        </thead>
        <tbody>
          {faqItems.map(item => (
            <tr key={item.Id}>
              <td style={{ borderBottom: '1px solid #eee' }}>{item.Title}</td>
              <td style={{ borderBottom: '1px solid #eee' }}>{item.Description}</td>
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
                  onClick={() => handleDelete({ Id: item.Id, Title: item.Title })}
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
