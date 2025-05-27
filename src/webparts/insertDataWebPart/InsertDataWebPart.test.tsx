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

// Provide all required props for IInsertDataWebPartProps
// Mock a minimal WebPartContext for testing
import { WebPartContext } from '@microsoft/sp-webpart-base';
// function selectDropdownOption(arg0: string, arg1: string) {
//   throw new Error('Function not implemented.');
// }
const mockContext = {} as unknown as WebPartContext;

const mockProps = {
  context: mockContext,
  description: '',
  isDarkTheme: false,
  environmentMessage: '',
  hasTeamsContext: false,
  userDisplayName: ''
};

describe('InsertDataWebPart', () => {
  it('renders Create Item button', () => {
    render(<InsertDataWebPart {...mockProps} />);
    expect(screen.getByText('Create Item')).toBeInTheDocument();
  });

  it('opens the form dialog when Create Item is clicked', () => {
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    expect(screen.getByText('Create FAQ Item')).toBeInTheDocument();
  });

  it('shows validation errors if required fields are empty after blur', async () => {
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    // Focus and blur Title
    const titleInput = screen.getByLabelText('Title');
    titleInput.focus();
    titleInput.blur();
    // Focus and blur Body
    const bodyInput = screen.getByLabelText('Body');
    bodyInput.focus();
    bodyInput.blur();
    // Focus and blur Letter (dropdown)
    const letterDropdown = screen.getByLabelText('Letter');
    letterDropdown.focus();
    letterDropdown.blur();
    // Fluent UI renders errors as aria-live messages, so query by role alert
    const alerts = await screen.findAllByRole('alert');
    expect(alerts.some(a => a.textContent?.match(/Title is required/i))).toBe(true);
    expect(alerts.some(a => a.textContent?.match(/Body is required/i))).toBe(true);
    expect(alerts.some(a => a.textContent?.match(/Letter is required/i))).toBe(true);
  });

  async function selectDropdownOption(label: string, optionText: string) {
    userEvent.click(screen.getByLabelText(label));
    // The dropdown options are rendered in a portal, so use getByRole('listbox')
    const listbox = await screen.findByRole('listbox');
    const option = within(listbox).getByText(optionText);
    userEvent.click(option);
  }

  it('create an item successfully', async () => {
    render(<InsertDataWebPart {...mockProps}/>);
    userEvent.click(screen.getByText('Create Item'));
    const fakeTitle = faker.lorem.sentence();
    userEvent.type(screen.getByLabelText('Title'), fakeTitle);
    userEvent.type(screen.getByLabelText('Body'), fakeTitle);
    await selectDropdownOption('Letter', 'A');
    userEvent.click(screen.getByText('Submit'));
    // Wait for the MessageBar to appear, then check for the success message
    // const messageBar = await screen.findByTestId('success-message', {}, { timeout: 5000 });
    // expect(messageBar).toBeInTheDocument();
    // expect(within(messageBar).getByText(/item created successfully/i)).toBeInTheDocument();
    expect(screen.getByText(fakeTitle)).toBeInTheDocument();
    expect(screen.getByText('A')).toBeInTheDocument(); 
  });

  it('shows error message if item creation fails', async () => {
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    userEvent.type(screen.getByLabelText('Title'), 'Test Title');
    userEvent.type(screen.getByLabelText('Body'), 'Test Body');
    await selectDropdownOption('Letter', 'A');
    userEvent.click(screen.getByText('Submit'));
    expect(await screen.findByText('Error creating item')).toBeInTheDocument();
  });

  it('edits an item successfully', async () => {  
    // Reset fetch mock from previous tests to ensure success
    window.fetch = jest.fn(() => 
      Promise.resolve({
        ok: true,
        json: () => Promise.resolve({ Id: 1, Title: 'Test Title', Body: 'Test Body', Letter: 'A' })
      })
    ) as jest.Mock;
    
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    const fakeTitle = faker.lorem.sentence();
    userEvent.type(screen.getByLabelText('Title'), fakeTitle);
    userEvent.type(screen.getByLabelText('Body'), fakeTitle);
    await selectDropdownOption('Letter', 'A');
    userEvent.click(screen.getByText('Submit'));
    userEvent.click(screen.getByText('Close'));
    
    // Wait for the form dialog to close and the item to appear in the list
    await screen.findByText(fakeTitle);

    // Find the item row by its title

    const itemRow = await screen.findByText(fakeTitle);
    expect(itemRow).toBeInTheDocument();

    const row = itemRow.closest('tr');
    // Use querySelector to find the edit button by data-testid prefix
    const editButton = row && row.querySelector('[data-testid^="edit-button-"]');
    expect(editButton).not.toBeNull();
    // userEvent.click requires a non-null Element, so cast is safe after the check
    userEvent.click(editButton as HTMLElement);
    
    // Perform the edit
    // const newFakeTitle = faker.lorem.sentence();
    // const titleInput = await screen.findByLabelText('Title');
    // userEvent.clear(titleInput);
    // userEvent.type(titleInput, newFakeTitle);
    // userEvent.click(screen.getByText('Update'));

    // // Verify the updated item appears
    // expect(await screen.findByText(newFakeTitle)).toBeInTheDocument();
  });

  it('deletes an item successfully', async () => {
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    const fakeTitle = faker.lorem.sentence();
    userEvent.type(screen.getByLabelText('Title'), fakeTitle);
    userEvent.type(screen.getByLabelText('Body'), faker.lorem.paragraph());
    await selectDropdownOption('Letter', 'A');
    userEvent.click(screen.getByText('Submit'));
    // Wait for the dialog to close and the item to appear in the list
    await screen.findByText(fakeTitle);
    // Now delete the item
    userEvent.click(screen.getByText('Delete'));
    expect(screen.queryByText(fakeTitle)).not.toBeInTheDocument();
  });

  it('shows error message if item deletion fails', async () => {
    window.fetch = jest.fn(() => Promise.reject(new Error('Network error'))) as jest.Mock;
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    userEvent.type(screen.getByLabelText('Title'), 'Test Title');
    userEvent.type(screen.getByLabelText('Body'), 'Test Body');
    await selectDropdownOption('Letter', 'A');
    userEvent.click(screen.getByText('Submit'));
    // Now try to delete the item
    userEvent.click(screen.getByText('Delete'));
    expect(await screen.findByText('Error deleting item')).toBeInTheDocument();
  });

  it('renders edit and delete buttons with correct data-testid and allows extracting item Id', async () => {
  render(<InsertDataWebPart {...mockProps} />);
  userEvent.click(screen.getByText('Create Item'));
  const fakeTitle = faker.lorem.sentence();
  userEvent.type(screen.getByLabelText('Title'), fakeTitle);
  userEvent.type(screen.getByLabelText('Body'), 'Test body');
  await selectDropdownOption('Letter', 'A');
  userEvent.click(screen.getByText('Submit'));

  // Find the edit button by data-testid pattern
  const editButton = await screen.findByTestId(/^edit-button-\d+$/);
  expect(editButton).toBeInTheDocument();

  // Extract the item Id from the data-testid attribute
  const dataTestId = editButton.getAttribute('data-testid');
  expect(dataTestId).toMatch(/^edit-button-\d+$/);
  const idMatch = dataTestId?.match(/^edit-button-(\d+)$/);
  expect(idMatch).not.toBeNull();
  const itemId = idMatch ? Number(idMatch[1]) : null;
  expect(typeof itemId).toBe('number');
  expect(itemId).toBeGreaterThan(0);

  // The delete button should have the same Id
  const deleteButton = screen.getByTestId(`delete-button-${itemId}`);
  expect(deleteButton).toBeInTheDocument();
});

});



