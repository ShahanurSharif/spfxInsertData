// jest.setup.js
// Mock window.alert to prevent jsdom errors and allow assertions
global.alert = jest.fn();

// Suppress Fluent UI icon registration warnings (if any)
jest.spyOn(console, 'error').mockImplementation((msg, ...args) => {
  if (
    typeof msg === 'string' &&
    (msg.includes('registerIcons') || msg.includes('Icon'))
  ) {
    return;
  }
  // Uncomment below to see other errors in test output
  // console.warn('console.error:', msg, ...args);
});
