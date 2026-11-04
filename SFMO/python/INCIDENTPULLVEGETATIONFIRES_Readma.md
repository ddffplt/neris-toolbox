IncidentPull_VegetationFires — Run instructions

This script fetches NERIS incidents filtered to vegetation/wildfire Outside Fire types and exports an Excel workbook.

Important notes
- The script is deliberately conservative for statewide runs: by default it will restrict enumeration to MA unless you pass --state or --national.
- To avoid being locked out by MFA timing, the script will pause right before triggering the MFA challenge so you can control when the one-time code is sent and enter it into the same running process.

How to run (recommended, interactive):
1. Open a local terminal (recommended) in this folder.
2. Run: python IncidentPull_VegetationFires.py --statewide --state MA --range-type "Occurred Last 30 Days"
   - You may also pass --entity-id FDxxxxx to run for a single department.
3. The script will prompt for username and password if not provided via --username / --password.
4. You will see a prompt: "About to trigger the MFA challenge. When you are ready to receive the code, press Enter to continue."
   - Press Enter when you're ready to receive the MFA email/SMS.
5. The script will trigger the MFA and then prompt you to paste the one-time code. Paste it and press Enter.
6. The script will continue and export the workbook when done.

Non-interactive usage (fragile):
- You can pass --mfa-code <code> to supply the one-time code non-interactively, but the code must be the one generated for that exact run and this approach is fragile because the code expires quickly.

If you want me to run the script from this environment, I can do so but we must either:
- Pipe the MFA code into the run immediately when it is generated (fast, fragile), or
- Use --mfa-code with a code generated after the challenge (also fragile). 

Recommended: run locally and interact directly to avoid timing issues.
