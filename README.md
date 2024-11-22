# Yahoo Finance Sectors Crawler

A Python-based data crawler that retrieves 'Day Return' values for 11 sectors and 145 industries from Yahoo Finance, organizes the data into an Excel file, and sends it via email. The GUI functionality has been deprecated in the latest version.
If you prefer using the GUI program, please check v2.2.1 in the Release tab.

## Features
- Fetches 'Day Return' values for 11 sectors and 145 industries from Yahoo Finance.
  - Saves data into an Excel file (`.xlsx`) with headers for page, sector names, and percentage changes.
  > **Note:** The file will not be saved locally when the workflow runs on GitHub Actions. It will be temporarily generated in the GitHub Actions environment and will be deleted once the workflow execution completes. If you run the code locally, the file will be saved in your local directory.
- Automatically sends the Excel file via email.

## Setup Instructions
### 1. Fork this Repository
Make sure to fork this repository from the **main branch** or check the **Releases section** to select the branch corresponding to the version you need.

### 2.Set up environment variables
This program requires 3 environment variables.
These values must be saved in your repository's GitHub **Settings** under **Secrets and variables**.

Follow these steps:

1. Go to **Settings** → **Secrets and variables** → **Actions** → **Secrets** → **Repository secrets**.
2. Click the **"New repository secret"** button.
3. Add the following three secrets:
- `EMAIL_ADDRESS`: Sender's email address.
- `EMAIL_PASSWORD`: Sender's App Password (not the regular email password).
- `RECIPIENT_EMAIL`: Receiver's email address.

#### **Example Setup with Gmail**
For Gmail, you must use an **App Password**, not your regular email password. Follow these steps to create an App Password:

1. Visit the [Google Support page](https://support.google.com/accounts/answer/185833) and go to the **"Create & use app passwords"** section.
2. Click the link **"Create and manage your app passwords"**.
3. Assign a name like **"crawler"** to the application and generate an App Password.
4. Google will provide a 16-character App Password. Use this password as the value for `EMAIL_PASSWORD`.

### 3. Set Up the Workflow
To automate the script and receive daily updates, configure the workflow file as follows:

1. Navigate to `.github/workflows/schedule-crawler.yml`.
2. Uncomment the block below:
```yaml
# schedule:
#   - cron: '30 20 * * *'
```
3. Add your desired time in **UTC format**.

The `cron` expression follows this structure:
```yaml
cron: 'minute hour day-of-month month day-of-week'
```
> Note: The workflow runs based on UTC time, so always adjust your local time to UTC.

In this repository, the program runs every day, so `day-of-month`, `month`, and `day-of-week` will be set as `* * *`, meaning it runs daily.

#### Example 1
If you want to receive the email at **17:30 German time**:

1. Convert 17:30 German time to UTC.  
- German time is UTC+1 (or UTC+2 during daylight saving time).  
- For standard time: 17:30 in German time equals **16:30 UTC**.  
2. Write the cron expression as:
```yaml
cron: '30 16 * * *'
```
#### **Example 2**
To run at **8:00 AM Greek time (UTC+2)**:

1. Convert 8:00 AM Greek time to UTC.  
- Greek time is UTC+2 (or UTC+3 during daylight saving time).  
- For standard time: 8:00 AM Greek time equals 6:00 UTC.
2. Write the cron expression as:
```yaml
cron: '0 6 * * *'
```


### 4. Run the Workflow Manually (Optional)
If you want to test the workflow without waiting for the scheduled `cron` time:
1. Go to the **Actions** tab in your GitHub repository.
2. Select the action named **"Schedule Python Crawler"**.
3. Click the **"Run workflow"** button to trigger the workflow immediately.
