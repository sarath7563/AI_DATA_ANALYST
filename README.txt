DATAWISE AI — STAGE 3.1 AUTOMATED TESTS

WHAT THIS DOES
These tests check the Stage 2 app's existing calculation/helper functions without starting the Streamlit interface. They do not modify app.py.

INSTALL
1. Back up your current app.py first.
2. Copy the tests folder from this package into the same project folder that contains app.py.
   Example:
      AI_DATA_ANALYST/
        app.py
        tests/
          conftest.py
          test_data_quality.py
          test_analysis.py
3. Open the VS Code terminal in AI_DATA_ANALYST.
4. Install pytest:
      python -m pip install pytest

RUN ALL TESTS
      python -m pytest -v

EXPECTED RESULT
If the app.py file is the Stage 2 version and the required packages are installed, you should see 11 passed tests.

IMPORTANT
- These tests read helper functions from your local app.py. They do not replace your app.py.
- If you see a message saying expected helper functions are missing, confirm you copied the Stage 2 app to app.py.
- If a test fails, copy the complete terminal error and send it back so the test or app can be corrected.
- This is Stage 3.1: calculation and data-quality testing. Export-file validation will be the next sub-stage.
