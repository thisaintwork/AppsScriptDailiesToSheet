Act as an Expert Google Apps Script Developer. I have an existing Apps Script function that I need you to clean up, modernize, and standardize. 

Do not change the core business logic or the intended outcome of the script. Your job is strictly to improve the code quality and inject my standard error handling.

### 1. My Existing Code
Here is the script I have written so far:

[PASTE YOUR EXISTING FUNCTION HERE]

### 2. Standard Error Handling (CRITICAL)
I have a standard way of handling errors. You MUST wrap the core logic of my function in the exact error-handling pattern shown below. Do not invent your own error handling. 

[PASTE YOUR STANDARD ERROR HANDLING HERE. Example:]
/*
try {
  // Original core logic goes here
} catch (error) {
  myCustomErrorHandler(error, arguments.callee.name);
  return null;
}
*/

### 3. Cleanup & Refactoring Rules
In addition to adding the error handling, please refactor the code according to these rules:
1.  **JSDoc Comments:** Add a professional JSDoc comment block at the top explaining what the function does, its parameters, and its return value based on my code.
2.  **Modernize Syntax:** Convert old syntax to modern ES6+. Change `var` to `const` or `let`. Use arrow functions and template literals (backticks) where appropriate.
3.  **Variable Naming:** If my variable names are vague (like `x`, `y`, `data1`), rename them to be clear, descriptive, and camelCase based on their context.
4.  **Comments:** Add brief, helpful inline comments explaining the steps of the code. 
5.  **Clean Formatting:** Ensure proper indentation and spacing.

Please provide only the complete, ready-to-use refactored Apps Script code block.