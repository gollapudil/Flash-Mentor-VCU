

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mentors"); // ← CONFIRM this matches Admin.js's adminSheet tab name

// Cache for column indexes to avoid repeated searches
var columnIndexCache = null;
var lastDataTimestamp = null;
var cachedMentorData = null;

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;
  var page   = e && e.parameter ? e.parameter.page   : null;

  if (action) {
    var result;
    try {
      if (action === "getMentors") {
        result = { mentors: getCachedMentorData() };

      } else if (action === "getAdminOverview") {
        // NEW: proper admin dashboard data (stats + mentors), from Admin.js
        result = getAdminOverview();

      } else if (action === "checkUserBookingStatus") {
        result = checkUserBookingStatus(e.parameter.email);

      } else if (action === "getMentorProfile") {
        result = getMentorProfile(e.parameter.mentorName, e.parameter.email);

      } else if (action === "bookSlot") {
        result = bookSlot(e.parameter.mentorName, e.parameter.email);

      } else if (action === "addStudentToMentor") {
        // FIXED: was using undefined "data" — now uses e.parameter (this is doGet, not doPost)
        result = addStudentToMentor(e.parameter.mentorName, e.parameter.email);

      } else if (action === "removeStudentFromMentor") {
        // FIXED: was using undefined "data" — now uses e.parameter
        result = removeStudentFromMentor(e.parameter.mentorName, e.parameter.email);

      } else {
        result = { error: "Unknown action" };
      }
    } catch (err) {
      result = { error: "Server error: " + err.message };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Fallback: still serve admin HTML if needed
  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle("VCU Admin Dashboard - Mentor Management")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default fallback (won't be used once Netlify is live)
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle("VCU Engineering Alternate Assignment Career Conversation")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  // Parse JSON body sent by callGAS() in index.html / admin.html
  var data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "Invalid JSON: " + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var action = data.action;
  var result;

  try {
    if (action === "sendVerificationCode") {
      result = sendVerificationCode(data.email);

    } else if (action === "verifyCodeAndAuthenticate") {
      result = verifyCodeAndAuthenticate(data.email, data.code);

    } else if (action === "getMentors") {
      result = { mentors: getCachedMentorData() };

    } else if (action === "getAdminOverview") {
      // NEW: proper admin dashboard data (stats + mentors), from Admin.js
      result = getAdminOverview();

    } else if (action === "checkUserBookingStatus") {
      result = checkUserBookingStatus(data.email);

    } else if (action === "getMentorProfile") {
      result = getMentorProfile(data.mentorName, data.email);

    } else if (action === "bookSlot") {
      result = bookSlot(data.mentor, data.email); // ← note: "mentor" not "mentorName"

    } else if (action === "markBookingDone") {
      result = markBookingDone(data);

    } else if (action === "addStudentToMentor") {
      result = addStudentToMentor(data.mentorName, data.email);

    } else if (action === "removeStudentFromMentor") {
      result = removeStudentFromMentor(data.mentorName, data.email);

    } else {
      result = { error: "Unknown action: " + action };
    }

  } catch (err) {
    result = { error: "Server error: " + err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Generate random verification code
function generateVerificationCode() {
  return Math.floor(100000 + Math.random() * 900000).toString(); // 6-digit code
}

// Validate VCU email format
function validateVCUEmail(email) {
  if (!email || typeof email !== 'string') {
    return { valid: false, message: "Please enter an email address." };
  }
  email = email.trim().toLowerCase();

  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    return { valid: false, message: "Please enter a valid email address." };
  }

  if (!email.endsWith("@vcu.edu")) {
    return { valid: false, message: "Only VCU students with @vcu.edu email addresses can access this system." };
  }

  return { valid: true, email: email };
}

// Send verification code to email
function sendVerificationCode(email) {
  try {
    var validation = validateVCUEmail(email);
    if (!validation.valid) {
      return { success: false, message: validation.message };
    }

    var validEmail = validation.email;
    var verificationCode = generateVerificationCode();

    var expirationTime = new Date().getTime() + (10 * 60 * 1000); // 10 minutes
    PropertiesService.getScriptProperties().setProperty(
      'verification_' + validEmail,
      JSON.stringify({
        code: verificationCode,
        expires: expirationTime
      })
    );

    var emailSubject = "VCU Engineering Alternate Assignment Career Conversation - Verification Code";
    var emailBody = "Dear VCU Student,\n\nYour verification code for VCU Engineering Alternate Assignment Career Conversation is:\n\n🔐 VERIFICATION CODE: " + verificationCode + "\n\nThis code will expire in 10 minutes. Please enter this code on the website to access the mentor booking system.\n\nIf you did not request this code, please ignore this email.\n\nBest regards,\nVCU College of Engineering Career Services Team\n\n---\nThis is an automated message. Please do not reply to this email.";

    try {
      // NOTE: to send from a department alias (e.g. engg@vcu.edu), add:
      // from: "engg@vcu.edu", name: "VCU Engineering Career Services"
      // — but only after that alias is verified under "Send mail as" in Gmail settings.
      MailApp.sendEmail({
        to: validEmail,
        subject: emailSubject,
        body: emailBody
      });

      return {
        success: true,
        email: validEmail,
        message: "Verification code sent! Check your email."
      };

    } catch (emailError) {
      console.error("Failed to send verification email:", emailError);
      return {
        success: false,
        message: "Failed to send verification email. Please check if your email address is correct."
      };
    }

  } catch (error) {
    console.error("Error in sendVerificationCode:", error);
    return {
      success: false,
      message: "Failed to send verification code. Please try again."
    };
  }
}

// Verify the code and authenticate user
function verifyCodeAndAuthenticate(email, code) {
  try {
    var validation = validateVCUEmail(email);
    if (!validation.valid) {
      return { success: false, message: validation.message };
    }

    var validEmail = validation.email;

    var verificationData = PropertiesService.getScriptProperties().getProperty('verification_' + validEmail);

    if (!verificationData) {
      return { success: false, message: "No verification code found. Please request a new code." };
    }

    var parsedData = JSON.parse(verificationData);
    var currentTime = new Date().getTime();

    if (currentTime > parsedData.expires) {
      PropertiesService.getScriptProperties().deleteProperty('verification_' + validEmail);
      return { success: false, message: "Verification code has expired. Please request a new code." };
    }

    if (code.trim() !== parsedData.code) {
      return { success: false, message: "Invalid verification code. Please check your email and try again." };
    }

    PropertiesService.getScriptProperties().deleteProperty('verification_' + validEmail);

    var sessionId = Utilities.getUuid();
    var sessionData = {
      email: validEmail,
      timestamp: new Date().getTime(),
      verified: true
    };

    PropertiesService.getScriptProperties().setProperty('session_' + sessionId, JSON.stringify(sessionData));

    return {
      success: true,
      email: validEmail,
      sessionId: sessionId,
      message: "Email verified successfully!"
    };

  } catch (error) {
    console.error("Error in verifyCodeAndAuthenticate:", error);
    return {
      success: false,
      message: "Verification failed. Please try again."
    };
  }
}

// Get and cache column indexes once
function getColumnIndexes() {
  if (columnIndexCache !== null) {
    return columnIndexCache;
  }
  try {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var indexes = {
      name: headers.indexOf("First Name & Last Name"),
      areaOfFocus: headers.indexOf("What is your area of focus?"),
      industry: headers.indexOf("Industry you can share about."),
      major: headers.indexOf("What was your major?"),
      additionalInfo: headers.indexOf("Any other information about yourself that might be helpful to a student in determining whom to talk with? "),
      linkedin: headers.indexOf("What is your LinkedIn Profile?"),
      email: headers.indexOf("Email Address"),
      preferredEmail: headers.indexOf("Preferred email address for students to use to contact you. This email will only be shared with the specific student(s) who sign up to talk with you. "),
      company: headers.indexOf("Company"),
      companyAlt: headers.indexOf("What company are you with or previously with?"),
      slots: headers.indexOf("Available Slots"),
      signedUp: headers.indexOf("Signed-Up Students"),
      conversations: -1
    };

    if (indexes.email === -1) {
      indexes.email = indexes.preferredEmail;
    }

    if (indexes.company === -1) {
      indexes.company = indexes.companyAlt;
    }

    for (var i = 0; i < headers.length; i++) {
      var header = headers[i].trim();
      if (header.includes("How many conversations") && header.includes("would you be open to having")) {
        indexes.conversations = i;
        break;
      }
    }

    if (indexes.slots === -1 && indexes.conversations !== -1) {
      indexes.slots = indexes.conversations;
    }

    if (indexes.signedUp === -1) {
      console.log("Adding 'Signed-Up Students' column...");
      var lastColumn = headers.length + 1;
      sheet.getRange(1, lastColumn).setValue("Signed-Up Students");
      indexes.signedUp = lastColumn - 1;
    }

    columnIndexCache = indexes;
    console.log("Column indexes cached:", indexes);
    return indexes;

  } catch (error) {
    console.error("Error getting column indexes:", error);
    throw new Error("Failed to map spreadsheet columns: " + error.message);
  }
}

// Get cached mentor data or read fresh data
function getCachedMentorData(forceRefresh) {
  if (forceRefresh === undefined) forceRefresh = false;

  try {
    var currentTimestamp = new Date().getTime();

    if (!forceRefresh && cachedMentorData && lastDataTimestamp &&
        (currentTimestamp - lastDataTimestamp) < 30000) {
      console.log("Using cached mentor data");
      return cachedMentorData;
    }

    console.log("Reading fresh mentor data from spreadsheet");

    var indexes = getColumnIndexes();
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();

    if (data.length <= 1) {
      cachedMentorData = [];
      lastDataTimestamp = currentTimestamp;
      return [];
    }

    var mentors = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // FIXED: signedUp column stores a comma-separated list of student emails
      // (that's what bookSlot/addStudentToMentor write). The old code did
      // parseInt() on this string, which always returned 0. Now we split it
      // into a real array, matching how bookSlot/checkUserBookingStatus use it.
      var signedUpRaw = row[indexes.signedUp] || "";
      var signedUpList = signedUpRaw
        ? signedUpRaw.toString().split(",").map(function(e) { return e.trim(); }).filter(function(e) { return e !== ""; })
        : [];

      mentors.push({
        name: row[indexes.name] || "",
        areaOfFocus: row[indexes.areaOfFocus] || "",
        industry: row[indexes.industry] || "",
        company: row[indexes.company] || "",
        major: row[indexes.major] || "",
        additionalInfo: row[indexes.additionalInfo] || "",
        linkedinUrl: row[indexes.linkedin] || "",
        email: row[indexes.email] || "",
        availableSlots: parseInt(row[indexes.slots]) || 0,
        signedUpStudents: signedUpList, // ← now a real array of emails, not a broken number
        rowIndex: i + 1
      });
    }

    cachedMentorData = mentors;
    lastDataTimestamp = currentTimestamp;
    console.log("Cached " + mentors.length + " mentors");

    return mentors;

  } catch (error) {
    console.error("Error getting mentor data:", error);
    throw new Error("Failed to load mentor data: " + error.message);
  }
}

// Clear cache when data is modified
function clearMentorDataCache() {
  cachedMentorData = null;
  lastDataTimestamp = null;
  console.log("Mentor data cache cleared");
}

// Get detailed mentor information
function getMentorProfile(mentorName, userEmail) {
  try {
    var studentEmail = null;
    if (userEmail) {
      var validation = validateVCUEmail(userEmail);
      if (!validation.valid) {
        throw new Error("Invalid email: " + validation.message);
      }
      studentEmail = validation.email;
    }

    var mentors = getCachedMentorData();

    var mentor = null;
    for (var i = 0; i < mentors.length; i++) {
      if (mentors[i].name === mentorName) {
        mentor = mentors[i];
        break;
      }
    }

    if (!mentor) {
      throw new Error("Mentor not found: " + mentorName);
    }

    var linkedinUrl = mentor.linkedinUrl;
    if (linkedinUrl && linkedinUrl.indexOf("http") !== 0) {
      linkedinUrl = "https://" + linkedinUrl;
    }

    return {
      name: mentor.name,
      areaOfFocus: mentor.areaOfFocus,
      industry: mentor.industry,
      company: mentor.company,
      major: mentor.major,
      additionalInfo: mentor.additionalInfo,
      linkedinUrl: linkedinUrl,
      email: mentor.email,
      availableSlots: mentor.availableSlots,
      signedUpStudents: mentor.signedUpStudents,
      isBookedByCurrentUser: studentEmail ? mentor.signedUpStudents.indexOf(studentEmail) !== -1 : false
    };

  } catch (error) {
    console.error("Error in getMentorProfile:", error);
    throw new Error("Failed to load mentor profile: " + error.message);
  }
}

// Book a slot
function bookSlot(mentorName, userEmail) {
  try {
    var validation = validateVCUEmail(userEmail);
    if (!validation.valid) {
      return { success: false, message: "Invalid email: " + validation.message };
    }
    var studentEmail = validation.email;
    var indexes = getColumnIndexes();

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][indexes.name] === mentorName) {
        var availableSlots = parseInt(data[i][indexes.slots]) || 0;

        if (availableSlots <= 0) {
          return { success: false, message: "This mentor has no available slots remaining." };
        }

        sheet.getRange(i + 1, indexes.slots + 1).setValue(availableSlots - 1);

        var currentStudents = data[i][indexes.signedUp] || "";
        var updatedStudents = currentStudents + (currentStudents ? ", " : "") + studentEmail;
        sheet.getRange(i + 1, indexes.signedUp + 1).setValue(updatedStudents);

        clearMentorDataCache();

        var mentorEmail = data[i][indexes.email];
        var areaOfFocus = data[i][indexes.areaOfFocus];
        var industry = data[i][indexes.industry];
        var company = data[i][indexes.company];

        var linkedinUrl = data[i][indexes.linkedin] || "";
        if (linkedinUrl && linkedinUrl.indexOf("http") !== 0) {
          linkedinUrl = "https://" + linkedinUrl;
        }

        var emailSubject = "VCU Engineering Alternate Assignment Career Conversation Mentor Confirmed - " + mentorName;
        var emailBody = "Dear " + studentEmail + ",\n\n" +
          "CONGRATULATIONS! You have selected " + mentorName + " as your mentor for career conversation.\n\n" +
          "Name: " + mentorName + "\n" +
          "Email: " + mentorEmail + "\n" +
          "Area of Focus: " + areaOfFocus + "\n" +
          "Industry: " + industry + "\n" +
          "Company: " + company + "\n" +
          (linkedinUrl ? "LinkedIn Profile: " + linkedinUrl + "\n" : "") +
          "\nWe're excited for you to connect with " + mentorName + "!\n\n" +
          "Best regards,\nVCU College of Engineering Career Services Team\n\n" +
          "This is an automated message. Please do not reply to this email.";

        try {
          GmailApp.sendEmail(  studentEmail, emailSubject, emailBody, {from: "engrstwk2@vcu.edu", name:"Engineering Career Services"} );
        } catch (emailError) {
          console.error("Email sending failed:", emailError); // don't fail the booking if email fails
        }

        return { success: true, message: "Successfully booked with " + mentorName + "! Check your VCU email " + studentEmail + " for confirmation details." };
      }
    }

    return { success: false, message: "Mentor not found or no longer available." };

  } catch (error) {
    console.error("Error in bookSlot:", error);
    return { success: false, message: "Booking failed: " + error.message };
  }
}

// Check if user has already booked a mentor
function checkUserBookingStatus(userEmail) {
  try {
    var validation = validateVCUEmail(userEmail);
    if (!validation.valid) {
      return { hasBooked: false, error: "Invalid email" };
    }
    var studentEmail = validation.email;

    var mentors = getCachedMentorData();
    for (var i = 0; i < mentors.length; i++) {
      if (mentors[i].signedUpStudents.indexOf(studentEmail) !== -1) {
        return { hasBooked: true, mentorName: mentors[i].name };
      }
    }
    return { hasBooked: false };

  } catch (error) {
    console.error("Error checking booking status:", error);
    return { hasBooked: false, error: error.message };
  }
}

function markBookingDone(data) {
  try {
    var mentorName = data.mentor;
    var studentEmail = data.email;
    if (!mentorName || !studentEmail) {
      return { success: false, message: "Missing mentor or email" };
    }

    var indexes = getColumnIndexes();
    var values = sheet.getDataRange().getValues();

    for (var i = 1; i < values.length; i++) {
      if (values[i][indexes.name] === mentorName) {
        var currentStudents = values[i][indexes.signedUp] ? values[i][indexes.signedUp].toString() : "";
        var updatedList = currentStudents
          .split(",")
          .map(function(s) { return s.trim(); })
          .filter(function(s) { return s !== "" && s !== studentEmail; })
          .join(", ");

        sheet.getRange(i + 1, indexes.signedUp + 1).setValue(updatedList);
        clearMentorDataCache();
        return { success: true, message: "Marked as done" };
      }
    }
    return { success: false, message: "Mentor not found: " + mentorName };

  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  }
}

// Debug function to check spreadsheet columns
function debugSpreadsheetColumns() {
  try {
    var indexes = getColumnIndexes();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    console.log("=== SPREADSHEET DEBUG INFO ===");
    console.log("Total columns found:", headers.length);
    console.log("Total rows found:", data.length);
    console.log("Column indexes:", indexes);
    headers.forEach(function(header, index) {
      console.log("Column " + index + ":", header);
    });
    return { totalColumns: headers.length, totalRows: data.length, headers: headers, indexes: indexes, success: true };
  } catch (error) {
    console.error("Debug error:", error);
    return { success: false, error: error.message };
  }
}
