import * as React from 'react';
import type { IInsertDataWebPartProps } from './IInsertDataWebPartProps';
import { Dropdown, IDropdownOption, TextField, PrimaryButton, MessageBar, MessageBarType } from '@fluentui/react';
import { sp } from '@pnp/sp-commonjs';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

const InsertDataWebPart: React.FC<IInsertDataWebPartProps> = (props) => {
  const [disabled, setDisabled] = React.useState(false);
  const [Title, setTitle] = React.useState('');
  const [Body, setBody] = React.useState('');
  const [Letter, setLetter] = React.useState('');
  const [options, setOptions] = React.useState<IDropdownOption[]>([]);
  const [successMessage, setSuccessMessage] = React.useState<string | null>(null);

  const [titleError, setTitleError] = React.useState<string | undefined>(undefined);
  const [bodyError, setBodyError] = React.useState<string | undefined>(undefined);
  const [letterError, setLetterError] = React.useState<string | undefined>(undefined);

  const validateTitle = (value?: string)=>{
    if (!value || value.trim() === '') {
      setTitleError('Title is required');
      return false;
    }
    setTitleError(undefined);
    return true;
  }

  const validateBody = (value?: string)=>{
    if (!value || value.trim() === '') {
      setBodyError('Body is required');
      return false;
    }
    setBodyError(undefined);
    return true;  
  }

  const validateLetter = (value?: string)=>{
    if (!value || value.trim() === '') {
      setLetterError('Letter is required');
      return false;
    }
    setLetterError(undefined);
    return true;  
  }

  React.useEffect(() => {
    // Disable if any field is empty or has an error
    setDisabled(
      !Title || !Body || !Letter || !!titleError || !!bodyError || !!letterError
    );
  }, [Title, Body, Letter, titleError, bodyError, letterError]);

  // Setup PnPjs for SPFx context
  React.useEffect(() => {
    sp.setup({ spfxContext: props.context as any });
  }, [props.context]);

  // Fetch dropdown options from the server
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

  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    // Validate all fields before submit
    const isTitleValid = validateTitle(Title);
    const isBodyValid = validateBody(Body);
    const isLetterValid = validateLetter(Letter);
    if (!isTitleValid || !isBodyValid || !isLetterValid) {
      alert('Please fill in all required fields correctly.');
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
