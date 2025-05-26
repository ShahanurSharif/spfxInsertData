/// <reference types="jest" />
import { render, screen, fireEvent } from '@testing-library/react';
import InsertDataWebPart from '../../src/webparts/insertDataWebPart/components/InsertDataWebPart';

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
});