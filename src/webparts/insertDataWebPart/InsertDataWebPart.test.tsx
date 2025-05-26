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

  it('shows validation errors if required fields are empty', async () => {
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    userEvent.click(screen.getByText('Submit'));
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
    userEvent.type(screen.getByLabelText('Body'), faker.lorem.paragraph());
    await selectDropdownOption('Letter', 'A');
    userEvent.click(screen.getByText('Submit'));
    expect(await screen.findByText('Item created successfully')).toBeInTheDocument();
    expect(screen.getByText(fakeTitle)).toBeInTheDocument();
    expect(screen.getByText('A')).toBeInTheDocument(); 
  });

  it('shows error message if item creation fails', async () => {
    window.fetch = jest.fn(() => Promise.reject(new Error('Network error'))) as jest.Mock;
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    userEvent.type(screen.getByLabelText('Title'), 'Test Title');
    userEvent.type(screen.getByLabelText('Body'), 'Test Body');
    await selectDropdownOption('Letter', 'A');
    userEvent.click(screen.getByText('Submit'));
    expect(await screen.findByText('Error creating item')).toBeInTheDocument();
  });

  it('edits an item successfully', async () => {  
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    const fakeTitle = faker.lorem.sentence();
    userEvent.type(screen.getByLabelText('Title'), fakeTitle);
    userEvent.type(screen.getByLabelText('Body'), faker.lorem.paragraph());
    await selectDropdownOption('Letter', 'A');
    userEvent.click(screen.getByText('Submit'));
    // Now edit the item
    userEvent.click(screen.getByText('Edit'));
    const newFakeTitle = faker.lorem.sentence();
    const titleInput = screen.getByLabelText('Title');
    userEvent.clear(titleInput);
    userEvent.type(titleInput, newFakeTitle);
    userEvent.click(screen.getByText('Submit'));
    expect(await screen.findByText('Item updated successfully')).toBeInTheDocument();
    expect(screen.getByText(newFakeTitle)).toBeInTheDocument();
  });

  it('deletes an item successfully', async () => {
    render(<InsertDataWebPart {...mockProps} />);
    userEvent.click(screen.getByText('Create Item'));
    const fakeTitle = faker.lorem.sentence();
    userEvent.type(screen.getByLabelText('Title'), fakeTitle);
    userEvent.type(screen.getByLabelText('Body'), faker.lorem.paragraph());
    await selectDropdownOption('Letter', 'A');
    userEvent.click(screen.getByText('Submit'));
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
});