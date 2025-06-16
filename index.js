const { chromium } = require("playwright");
const { humanKeys, humanSleep } = require("./utils");
const XLSX = require("xlsx");
const fs = require("fs");

const { jobTitles, payPeriodMap, states } = require("./constant");

function getEmailFromExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet);
  // Assume the first row has the email field
  return data[0]?.email || "test@example.com";
}

async function waitForAnySelector(page, selectors, options = {}) {
  for (let attempt = 0; attempt < (options.retries || 3); attempt++) {
    for (const selector of selectors) {
      try {
        const el = await page.waitForSelector(selector, {
          ...options,
          timeout: options.timeout || 5000,
        });
        if (el) {
          console.log(`[waitForAnySelector] Found: ${selector}`);
          return el;
        }
      } catch (e) {
        // Try next selector
      }
    }
    await humanSleep(1);
  }
  throw new Error(
    `[waitForAnySelector] None of the selectors found: ${selectors.join(", ")}`
  );
}

async function saveErrorState(page, prefix = "error") {
  try {
    await page.screenshot({ path: `${prefix}-screenshot.png` });
    const html = await page.content();
    fs.writeFileSync(`${prefix}-page.html`, html);
    console.log(
      `[saveErrorState] Saved screenshot and HTML as ${prefix}-screenshot.png and ${prefix}-page.html`
    );
  } catch (e) {
    console.error("[saveErrorState] Failed:", e);
  }
}

async function retryStep(fn, retries = 3, stepName = "step") {
  for (let i = 0; i < retries; i++) {
    try {
      return await fn();
    } catch (e) {
      console.warn(
        `[retryStep] ${stepName} failed (attempt ${i + 1}/${retries}):`,
        e.message
      );
      if (i === retries - 1) throw e;
      await humanSleep(2);
    }
  }
}

function getRandomDOB() {
  const start = new Date(1945, 0, 1).getTime();
  const end = new Date(2006, 11, 31).getTime();
  const dob = new Date(start + Math.random() * (end - start));
  // Format as mm/dd/yyyy with leading zeros
  const mm = String(dob.getMonth() + 1).padStart(2, "0");
  const dd = String(dob.getDate()).padStart(2, "0");
  const yyyy = dob.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
}

function excelDateToParts(excelDate) {
  const jsDate = new Date((excelDate - 25569) * 86400 * 1000);
  const mm = String(jsDate.getMonth() + 1).padStart(2, "0");
  const dd = String(jsDate.getDate()).padStart(2, "0");
  const yyyy = jsDate.getFullYear();
  return [mm, dd, yyyy];
}

function formatPayDate(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return [mm, dd, yyyy];
}

async function scrapeWebsite() {
  let browser;
  let page;

  try {
    // Launch the browser
    browser = await chromium.launch({
      headless: false,
      timeout: 60000,
    });

    // Create a new page
    page = await browser.newPage();
    page.setDefaultNavigationTimeout(60000);

    console.log("Attempting to navigate to the website...");
    await page.goto("https://offer.vivapaydayloans.com/?aff90131", {
      waitUntil: "networkidle",
      timeout: 60000,
    });
    console.log("Successfully loaded the website");

    // Wait for the apply button
    const applyBtn = await retryStep(
      () =>
        waitForAnySelector(page, ["#home-apply-button"], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for apply button"
    );
    await humanSleep(1);
    console.log("Clicking the apply button...");
    await applyBtn.click({ force: true });

    // Wait for the iframe to appear
    const iframeElement = await retryStep(
      () =>
        waitForAnySelector(page, ["#loan-application", "iframe.t-123"], {
          state: "attached",
          timeout: 15000,
        }),
      3,
      "wait for loan application iframe"
    );
    await humanSleep(1);

    // Get the iframe's content frame
    const frame = await iframeElement.contentFrame();
    if (!frame) throw new Error("Could not get iframe content frame");

    // Now, do everything inside the iframe!
    const email = getEmailFromExcel("data.xlsx");
    console.log("Filling email:", email);

    const emailInput = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#email", 'input[name="email"]', 'input[placeholder*="Email"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for email input in iframe"
    );
    await humanSleep(1);
    // Clear the input before typing
    await emailInput.fill("");
    await humanSleep(0.5);
    // Use humanKeys for human-like typing
    await humanKeys(emailInput, email);
    await humanSleep();
    // Fallback: set value directly if field is still empty (for stubborn fields)
    const currentValue = await frame.evaluate((el) => el.value, emailInput);
    if (!currentValue) {
      await frame.evaluate(
        (el, value) => {
          el.value = value;
          el.dispatchEvent(new Event("input", { bubbles: true }));
          el.dispatchEvent(new Event("change", { bubbles: true }));
        },
        emailInput,
        email
      );
      await humanSleep(0.5);
    }

    // Wait 3 seconds for the next fields to appear
    await humanSleep(3);

    // Read more data from Excel
    const data = XLSX.utils.sheet_to_json(
      XLSX.readFile("data.xlsx").Sheets[
        XLSX.readFile("data.xlsx").SheetNames[0]
      ]
    )[0];
    const firstName = data?.firstname || "John";
    const lastName = data?.lasname || "Doe";
    console.log("DOB:", data?.dob);
    let dobParts;
    if (typeof data?.dob === "number") {
      dobParts = excelDateToParts(data.dob);
    } else if (data?.dob) {
      // Assume string in MM/DD/YYYY or M/D/YYYY
      const [mm, dd, yyyy] = String(data.dob).split(/[\/\-]/);
      dobParts = [
        String(mm).padStart(2, "0"),
        String(dd).padStart(2, "0"),
        String(yyyy).padStart(4, "0"),
      ];
    } else {
      const rand = getRandomDOB().split("/");
      dobParts = [rand[0], rand[1], rand[2]];
    }
    const phone = data?.phone1 || "3233233323";

    // Fill First Name
    const firstNameInput = await retryStep(
      () =>
        waitForAnySelector(frame, ["#first_name", 'input[name="first_name"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for first name input in iframe"
    );
    await humanSleep(1);
    await firstNameInput.fill("");
    await humanKeys(firstNameInput, firstName);
    await humanSleep();

    // Fill Last Name
    const lastNameInput = await retryStep(
      () =>
        waitForAnySelector(frame, ["#last_name", 'input[name="last_name"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for last name input in iframe"
    );
    await humanSleep(1);
    await lastNameInput.fill("");
    await humanKeys(lastNameInput, lastName);
    await humanSleep();

    // Fill DOB
    const dobInput = await retryStep(
      () =>
        waitForAnySelector(frame, ["#dob", 'input[name="dob"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for dob input in iframe"
    );
    await humanSleep(1);
    await dobInput.fill(""); // Clear field
    await humanSleep(0.5);
    for (const part of dobParts) {
      await humanKeys(dobInput, part);
      await humanSleep(1);
    }

    // Fill Mobile Phone
    const phoneInput = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#mobile_phone", 'input[name="mobile_phone"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for mobile phone input in iframe"
    );
    await humanSleep(1);
    await phoneInput.fill("");
    await humanKeys(phoneInput, phone);
    await humanSleep();

    // Select 'Good' for credit score
    const creditScoreSelect = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#credit_score", 'select[name="credit_score"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for credit score select in iframe"
    );
    await humanSleep(1);
    await creditScoreSelect.selectOption({ value: "660" }); // Good (660-719)
    await humanSleep();

    // Select 'Yes' for checking account
    const checkAccount = await retryStep(() =>
      waitForAnySelector(
        frame,
        ["#active_checking_account", 'select[name="active_checking_account"]'],
        {
          state: "visible",
          timeout: 10000,
        }
      )
    );
    await humanSleep(1);
    await checkAccount.selectOption({ value: "1" });
    await humanSleep();

    // Fill address, zip, city, and state
    const address = data?.address || "123 Main St";
    const postCode = data?.zip || "90001";
    const city = data?.city || "Los Angeles";
    const stateAbbr = data?.state || "CA";
    const stateFull = states[stateAbbr.toUpperCase()] || "California";

    // Address
    const addressInput = await retryStep(
      () =>
        waitForAnySelector(frame, ["#address", 'input[name="address"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for address input in iframe"
    );
    await humanSleep(1);
    await addressInput.fill("");
    await humanKeys(addressInput, address);
    await humanSleep();

    // ZIP Code
    const zipInput = await retryStep(
      () =>
        waitForAnySelector(frame, ["#post_code", 'input[name="post_code"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for zip input in iframe"
    );
    await humanSleep(1);
    await zipInput.fill("");
    await humanKeys(zipInput, postCode);
    await humanSleep();

    // City
    const cityInput = await retryStep(
      () =>
        waitForAnySelector(frame, ["#city", 'input[name="city"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for city input in iframe"
    );
    await humanSleep(1);
    await cityInput.fill("");
    await humanKeys(cityInput, city);
    await humanSleep();

    // State
    const stateSelect = await retryStep(
      () =>
        waitForAnySelector(frame, ["#state", 'select[name="state"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for state select in iframe"
    );
    await humanSleep(1);
    await stateSelect.selectOption({ label: stateFull });
    await humanSleep();

    // Select 'Full Time Employed' for income type
    const incomeSourceSelect = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#income_source", 'select[name="income_source"]'],
          {
            state: "visible",
            timeout: 10000,
          }
        ),
      3,
      "wait for income source select in iframe"
    );
    await humanSleep(1);
    await incomeSourceSelect.selectOption({ label: "Full Time Employed" });
    await humanSleep();

    // Fill company name from employer in Excel
    const companyName = data?.employer || "Employer Inc";
    const companyInput = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#company_name", 'input[name="company_name"]'],
          {
            state: "visible",
            timeout: 10000,
          }
        ),
      3,
      "wait for company name input in iframe"
    );
    await humanSleep(1);
    await companyInput.fill("");
    await humanKeys(companyInput, companyName);
    await humanSleep();

    // Fill job title with a random job from the provided list
    const jobTitle = jobTitles[Math.floor(Math.random() * jobTitles.length)];
    const jobTitleInput = await retryStep(
      () =>
        waitForAnySelector(frame, ["#job_title", 'input[name="job_title"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for job title input in iframe"
    );
    await humanSleep(1);
    await jobTitleInput.fill("");
    await humanKeys(jobTitleInput, jobTitle);
    await humanSleep();

    const salary = data?.salary || "5000";
    const salaryInput = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#monthly_income", 'input[name="monthly_income"]'],
          {
            state: "visible",
            timeout: 10000,
          }
        ),
      3,
      "wait for salary input in iframe"
    );
    await humanSleep(1);
    await salaryInput.fill("");
    await humanKeys(salaryInput, salary);
    await humanSleep();

    // Fill work phone from workphone in Excel
    const workPhone = data?.workphone || "3233233323";
    const workPhoneInput = await retryStep(
      () =>
        waitForAnySelector(frame, ["#work_phone", 'input[name="work_phone"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for work phone input in iframe"
    );
    await humanSleep(1);
    await workPhoneInput.fill("");
    await humanKeys(workPhoneInput, workPhone);
    await humanSleep();

    // Select 'Direct Deposit' for income_payment_type
    const paymentTypeSelect = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#income_payment_type", 'select[name="income_payment_type"]'],
          {
            state: "visible",
            timeout: 10000,
          }
        ),
      3,
      "wait for income payment type select in iframe"
    );
    await humanSleep(1);
    await paymentTypeSelect.selectOption({ label: "Direct Deposit" });
    await humanSleep();

    // Select pay frequency from pay_period in Excel
    const payPeriodCode = data?.pay_period || "bi_weekly";
    const payFrequencyLabel = payPeriodMap[payPeriodCode] || "Weekly";
    const payFrequencySelect = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#pay_frequency", 'select[name="pay_frequency"]'],
          {
            state: "visible",
            timeout: 10000,
          }
        ),
      3,
      "wait for pay frequency select in iframe"
    );
    await humanSleep(1);
    await payFrequencySelect.selectOption({ label: payFrequencyLabel });
    await humanSleep();

    // Calculate next and following pay dates
    const today = new Date();
    const nextPayDate = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
    const followingPayDate = new Date(
      today.getTime() + 14 * 24 * 60 * 60 * 1000
    );

    const nextPayDateStr = formatPayDate(nextPayDate);
    const followingPayDateStr = formatPayDate(followingPayDate);

    // Fill Next Pay date
    const nextPayInput = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#next_payday", 'input[name="next_payday"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for next pay date input in iframe"
    );
    await humanSleep(1);
    await nextPayInput.fill("");
    await humanSleep(0.5);
    for (const part of nextPayDateStr) {
      await humanKeys(nextPayInput, part);
      await humanSleep(1);
    }
    await humanSleep();

    // Fill Following Pay date
    const followingPayInput = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#second_payday", 'input[name="second_payday"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for following pay date input in iframe"
    );
    await humanSleep(1);
    await followingPayInput.fill("");
    await humanSleep(0.5);
    for (const part of followingPayDateStr) {
      await humanKeys(followingPayInput, part);
      await humanSleep(1);
    }
    await humanSleep();

    // Bank Name
    const bankName = data?.bank_name || "Bank of America";
    const bankNameInput = await retryStep(
      () =>
        waitForAnySelector(frame, ['input[name="bank_name"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for bank name input in iframe"
    );
    await humanSleep(1);
    await bankNameInput.fill("");
    await humanKeys(bankNameInput, bankName);
    await humanSleep();

    // Bank Account Number
    const bankAccount = data?.bank_acc_number || "123456789";
    const bankAccountInput = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#bank_account_number", 'input[name="bank_account"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for bank account number input in iframe"
    );
    await humanSleep(1);
    await bankAccountInput.fill("");
    await humanKeys(bankAccountInput, bankAccount);
    await humanSleep();

    // Bank ABA
    let bankABA = data?.bank_routing || "011000015";
    bankABA = String(bankABA);
    if (bankABA.length < 9) {
      bankABA = bankABA.padEnd(9, "0");
    }
    const bankABAInput = await retryStep(
      () =>
        waitForAnySelector(frame, ["#bank_aba", 'input[name="bank_aba"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for bank ABA input in iframe"
    );
    await humanSleep(1);
    await bankABAInput.fill("");
    await humanKeys(bankABAInput, bankABA);
    await humanSleep();

    // Bank Type
    const bankType = data?.bank_acc_type === "checking" ? "Cheque" : "Savings";
    const bankTypeSelect = await retryStep(
      () =>
        waitForAnySelector(frame, ["#bank_type", 'select[name="bank_type"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for bank type select in iframe"
    );
    await humanSleep(1);
    await bankTypeSelect.selectOption({ value: bankType });
    await humanSleep();

    // Social Security Number
    const ssn = data?.ssn || "123456789";
    const ssnInput = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#social_security_number", 'input[name="social_security_number"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for ssn input in iframe"
    );
    await humanSleep(1);
    await ssnInput.fill("");
    await humanKeys(ssnInput, ssn);
    await humanSleep();

    // Driver License Number
    const dlNumber = data?.dl_number || "D1234567";
    const dlInput = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#driver_license_number", 'input[name="driver_license_number"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for driver license number input in iframe"
    );
    await humanSleep(1);
    await dlInput.fill("");
    await humanKeys(dlInput, dlNumber);
    await humanSleep();

    // Own Vehicle Free & Clear (should be 1)
    const ownVehicleSelect = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#own_vehicle_free_clear", 'select[name="own_vehicle_free_clear"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for own vehicle select in iframe"
    );
    await humanSleep(1);
    await ownVehicleSelect.selectOption({ value: "1" });
    await humanSleep();

    // In Debt (should be 0)
    const inDebtSelect = await retryStep(
      () =>
        waitForAnySelector(frame, ["#in_debt", 'select[name="in_debt"]'], {
          state: "visible",
          timeout: 10000,
        }),
      3,
      "wait for in debt select in iframe"
    );
    await humanSleep(1);
    await inDebtSelect.selectOption({ value: "0" });
    await humanSleep();

    // Check the credit report checkbox if present
    const creditReportCheckbox = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#credit_report_checkbox", 'input[name="credit_report_checkbox"]'],
          { state: "attached", timeout: 5000 }
        ),
      1,
      "wait for credit report checkbox in iframe"
    ).catch(() => null);
    if (creditReportCheckbox && !(await creditReportCheckbox.isChecked())) {
      await creditReportCheckbox.check();
      await humanSleep(0.5);
    }

    // Check the terms checkbox
    const termsCheckbox = await retryStep(
      () =>
        waitForAnySelector(frame, ['input[name="terms"]'], {
          state: "attached",
          timeout: 5000,
        }),
      1,
      "wait for terms checkbox in iframe"
    ).catch(() => null);
    if (termsCheckbox && !(await termsCheckbox.isChecked())) {
      await termsCheckbox.check();
      await humanSleep(0.5);
    }

    // Check the marketing consent checkbox
    const marketingCheckbox = await retryStep(
      () =>
        waitForAnySelector(frame, ["#term_email", 'input[name="term_email"]'], {
          state: "attached",
          timeout: 5000,
        }),
      1,
      "wait for marketing checkbox in iframe"
    ).catch(() => null);
    if (marketingCheckbox && !(await marketingCheckbox.isChecked())) {
      await marketingCheckbox.check();
      await humanSleep(0.5);
    }

    // Click the submit button
    const submitButton = await retryStep(
      () =>
        waitForAnySelector(
          frame,
          ["#register-submit", 'input[type="submit"]'],
          { state: "visible", timeout: 10000 }
        ),
      3,
      "wait for submit button in iframe"
    );
    await humanSleep(1);
    await submitButton.click();
    await humanSleep(2);

    console.log("Form submitted!");
  } catch (error) {
    console.error("An error occurred:", error);
    if (page) await saveErrorState(page, "error");
  } finally {
    if (browser) {
      await humanSleep(10); // Wait 10 seconds before closing the browser
      await browser.close();
      console.log("Browser closed");
    }
  }
}

scrapeWebsite();
