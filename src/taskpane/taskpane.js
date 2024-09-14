/* global console, document, Office */

Office.onReady((info) => {
  console.log("Office.onReady called with info:", info);
  if (info.host === Office.HostType.Excel) {
    console.log("Host is Excel. Proceeding to display app body.");
    // Hide the sideload message and display the app body
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Bind the execute-script button to its function
    const executeScriptButton = document.getElementById("execute-script");
    if (executeScriptButton) {
      executeScriptButton.onclick = executeScript;
    } else {
      console.warn("'execute-script' button not found in the DOM.");
    }
  } else {
    console.log("Host is not Excel. Add-in not initialized.");
  }
});

/**
 * Executes the Python script entered by the user, generates a plot, and displays it in the add-in.
 */
async function executeScript() {
  try {
    const scriptInput = document.getElementById("script-input").value;

    if (!scriptInput.trim()) {
      displayErrorDialog("Please enter a Python script to execute.");
      return;
    }

    console.log("Sending script to Python server.");

    // Show loading indicator
    const loadingIndicator = document.getElementById("loading-indicator");
    if (loadingIndicator) {
      loadingIndicator.style.display = "block";
    }

    const response = await fetch("http://127.0.0.1:8000/execute-script", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ script: scriptInput }),
    });

    console.log("Received response status:", response.status);

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(`Server error: ${errorData.detail}`);
    }

    const result = await response.json();

    console.log("Received data:", result);

    // Hide loading indicator
    if (loadingIndicator) {
      loadingIndicator.style.display = "none";
    }

    // Display the plot in the add-in
    displayPlot(result.plot);

    // Display any script output
    if (result.output) {
      displayOutput(result.output);
    }
  } catch (error) {
    console.error("Error executing script:", error);
    // Hide loading indicator in case of error
    const loadingIndicator = document.getElementById("loading-indicator");
    if (loadingIndicator) {
      loadingIndicator.style.display = "none";
    }
    displayErrorDialog("Failed to execute the Python script.\n" + error.message);
  }
}

/**
 * Displays the script output in the add-in.
 * @param {string} output - The output string from the script.
 */
function displayOutput(output) {
  const outputContainer = document.getElementById("script-output");
  if (outputContainer) {
    outputContainer.textContent = output;
    console.log("Script output displayed.");
  } else {
    console.error("'script-output' element not found in the DOM.");
    displayErrorDialog("Failed to display the script output.");
  }
}

/**
 * Displays the plot image received from the Python backend in the add-in.
 * @param {string} imgBase64 - The base64-encoded image string.
 */
function displayPlot(imgBase64) {
  const plotImage = document.getElementById("plot-image");
  if (plotImage) {
    console.log("Setting plot image with data length:", imgBase64.length);
    plotImage.src = `data:image/png;base64,${imgBase64}`;
    console.log("Plot image src set.");
  } else {
    console.error("'plot-image' element not found in the DOM.");
    displayErrorDialog("Failed to display the plot image.");
  }
}


/**
 * Displays an error dialog to the user with a specified message.
 * @param {string} message - The error message to display.
 */
function displayErrorDialog(message) {
  // Create the HTML content for the error dialog
  const errorHtml = `
    <!DOCTYPE html>
    <html>
      <head>
        <title>Error</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .error { color: red; }
          button { margin-top: 20px; padding: 10px 20px; }
        </style>
      </head>
      <body>
        <h2 class="error">Error</h2>
        <p>${message}</p>
        <button id="close-button">Close</button>

        <script>
          Office.initialize = function () {
            document.getElementById('close-button').onclick = function() {
              Office.context.ui.messageParent('Dialog closed');
            };
          };
        </script>
      </body>
    </html>
  `;

  // Open the error dialog
  Office.context.ui.displayDialogAsync(
    `data:text/html,${encodeURIComponent(errorHtml)}`,
    { height: 30, width: 20 },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to display error dialog:", asyncResult.error.message);
      }
    }
  );
}
