/// <reference types="jest" />
import { render, screen, fireEvent } from '@testing-library/react';
import '@testing-library/jest-dom';
import InsertDataWebPart from './components/InsertDataWebPart';
import { faker } from '@faker-js/faker';
import * as React from 'react';

// Provide all required props for IInsertDataWebPartProps
const mockProps = {
  context: {} as any,
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
    fireEvent.click(screen.getByText('Create Item'));
    expect(screen.getByText('Create FAQ Item')).toBeInTheDocument();
  });

  it('shows validation errors if required fields are empty', () => {
    render(<InsertDataWebPart {...mockProps} />);
    fireEvent.click(screen.getByText('Create Item'));
    fireEvent.click(screen.getByText('Submit'));
    expect(screen.getByText('Title is required')).toBeInTheDocument();
    expect(screen.getByText('Body is required')).toBeInTheDocument();
    expect(screen.getByText('Letter is required')).toBeInTheDocument();
  });

  // Add more tests for edit, delete, and form submission as needed
  it('create an item successfully', () => {
    render(<InsertDataWebPart {...mockProps}/>);
    fireEvent.click(screen.getByText('Create Item'));
    // Generate a fake title using faker
    const fakeTitle = faker.lorem.sentence();
    fireEvent.change(screen.getByLabelText('Title'), { target: { value: fakeTitle } });
    fireEvent.change(screen.getByLabelText('Body'), { target: { value: faker.lorem.paragraph() } });
    const letters = ['A', 'B', 'C'];
    const randomLetter = letters[Math.floor(Math.random() * letters.length)];
    fireEvent.change(screen.getByLabelText('Letter'), { target: { value: randomLetter } });
    fireEvent.click(screen.getByText('Submit'));
    expect(screen.getByText('Item created successfully')).toBeInTheDocument();
    expect(screen.getByText(fakeTitle)).toBeInTheDocument();
    expect(screen.getByText(randomLetter)).toBeInTheDocument(); 
  });

  it('shows error message if item creation fails', () => {
    // Mock the fetch call to simulate an error
    window.fetch = jest.fn(() =>
      Promise.reject(new Error('Network error'))
    ) as jest.Mock;

    render(<InsertDataWebPart {...mockProps} />);
    fireEvent.click(screen.getByText('Create Item'));
    fireEvent.change(screen.getByLabelText('Title'), { target: { value: 'Test Title' } });
    fireEvent.change(screen.getByLabelText('Body'), { target: { value: 'Test Body' } });
    fireEvent.change(screen.getByLabelText('Letter'), { target: { value: 'A' } });
    fireEvent.click(screen.getByText('Submit'));
    
    expect(screen.getByText('Error creating item')).toBeInTheDocument();
  });
  it('edits an item successfully', () => {  
    render(<InsertDataWebPart {...mockProps} />);
    fireEvent.click(screen.getByText('Create Item'));
    const fakeTitle = faker.lorem.sentence();
    fireEvent.change(screen.getByLabelText('Title'), { target: { value: fakeTitle } });
    fireEvent.change(screen.getByLabelText('Body'), { target: { value: faker.lorem.paragraph() } });
    const letters = ['A', 'B', 'C'];
    const randomLetter = letters[Math.floor(Math.random() * letters.length)];
    fireEvent.change(screen.getByLabelText('Letter'), { target: { value: randomLetter } });
    fireEvent.click(screen.getByText('Submit'));
    
    // Now edit the item
    fireEvent.click(screen.getByText('Edit'));
    const newFakeTitle = faker.lorem.sentence();
    fireEvent.change(screen.getByLabelText('Title'), { target: { value: newFakeTitle } });
    fireEvent.click(screen.getByText('Submit'));
    
    expect(screen.getByText('Item updated successfully')).toBeInTheDocument();
    expect(screen.getByText(newFakeTitle)).toBeInTheDocument();
  }
  );
  it('deletes an item successfully', () => {
    render(<InsertDataWebPart {...mockProps} />);
    fireEvent.click(screen.getByText('Create Item'));
    const fakeTitle = faker.lorem.sentence();
    fireEvent.change(screen.getByLabelText('Title'), { target: { value: fakeTitle } });
    fireEvent.change(screen.getByLabelText('Body'), { target: { value: faker.lorem.paragraph() } });
    const letters = ['A', 'B', 'C'];
    const randomLetter = letters[Math.floor(Math.random() * letters.length)];
    fireEvent.change(screen.getByLabelText('Letter'), { target: { value: randomLetter } });
    fireEvent.click(screen.getByText('Submit'));
    
    // Now delete the item
    fireEvent.click(screen.getByText('Delete'));
    
    expect(screen.queryByText(fakeTitle)).not.toBeInTheDocument();
  }
  );
  it('shows error message if item deletion fails', () => {
    // Mock the fetch call to simulate an error
    window.fetch = jest.fn(() =>
      Promise.reject(new Error('Network error'))
    ) as jest.Mock;

    render(<InsertDataWebPart {...mockProps} />);
    fireEvent.click(screen.getByText('Create Item'));
    fireEvent.change(screen.getByLabelText('Title'), { target: { value: 'Test Title' } });
    fireEvent.change(screen.getByLabelText('Body'), { target: { value: 'Test Body' } });
    fireEvent.change(screen.getByLabelText('Letter'), { target: { value: 'A' } });
    fireEvent.click(screen.getByText('Submit'));
    
    // Now try to delete the item
    fireEvent.click(screen.getByText('Delete'));

    expect(screen.getByText('Error deleting item')).toBeInTheDocument();
  });

});