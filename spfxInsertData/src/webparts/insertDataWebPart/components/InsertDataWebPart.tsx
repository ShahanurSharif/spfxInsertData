import * as React from 'react';
import type { IInsertDataWebPartProps } from './IInsertDataWebPartProps';
import { Dropdown, IDropdownOption, TextField, PrimaryButton, MessageBar, MessageBarType } from '@fluentui/react';
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

  // When you click the button, try to add the item
  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    const isTitleValid = validateTitle(Title);
    const isBodyValid = validateBody(Body);
    const isLetterValid = validateLetter(Letter);
    if (!isTitleValid || !isBodyValid || !isLetterValid) {
      return;
    }
    try {
      await sp.web.lists.getByTitle('FAQ').items.add({
        Title,
        body: Body,
        Letter
      });
      setSuccessMessage('FAQ item added successfully!');
      setTitle('');
      setBody('');
      setLetter('');
      setTimeout(() => setSuccessMessage(null), 5000);
    } catch (error) {
      alert('Error adding FAQ item: ' + error);
    }
  };

  // The form you see on the page
  return (
    <form onSubmit={handleSubmit}>
      {successMessage && (
        <MessageBar messageBarType={MessageBarType.success} isMultiline={false} onDismiss={() => setSuccessMessage(null)}>
          {successMessage}
        </MessageBar>
      )}
      <TextField 
        label='Title' 
        id='Title' 
        value={Title} 
        onChange={(_, v) => {
          setTitle(v || '');
          validateTitle(v);
        }} 
        onBlur={() => validateTitle(Title)}
        errorMessage={titleError}
        required
      />
      <TextField 
        label='Body' 
        id='Body' 
        value={Body} 
        onChange={(_, v) => {
          setBody(v || '');
          validateBody(v);
        }} 
        onBlur={() => validateBody(Body)}
        errorMessage={bodyError}
        multiline
        required
      />
      <Dropdown 
        label="Letter" 
        id="Letter" 
        options={options} 
        selectedKey={Letter} 
        onChange={(_, option) => {
          setLetter(option ? String(option.key) : '');
          validateLetter(option ? String(option.key) : '');
        }} 
        onBlur={() => validateLetter(Letter)}
        errorMessage={letterError}
        required
      />
      <br />
      <PrimaryButton text="Submit" type='submit' disabled={disabled} />
    </form>
  );
};

export default InsertDataWebPart;
