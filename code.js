var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");

// Cache for column indexes to avoid repeated searches
var columnIndexCache = null;
var lastDataTimestamp = null;
var cachedMentorData = null;

//different starting point if it's admin or students
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;
  var page   = e && e.parameter ? e.parameter.page   : null;

  // API calls from Netlify frontend
  if (action) {
    var result;

    if (action === "getMentors") {
      result = getCachedMentorData(e.parameter.email);
    } else if (action === "checkUserBookingStatus") {
      result = checkUserBookingStatus(e.parameter.email);
    } else if (action === "getMentorProfile") {
      result = getMentorProfile(e.parameter.mentorName, e.parameter.email);
    } else if (action === "bookSlot") {
      result = bookSlot(e.parameter.mentorName, e.parameter.email);
    } else {
      result = { error: "Unknown action" };
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
  // Parse JSON body sent by callGAS() in index.html
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
      result = getCachedMentorData(data.email);

    } else if (action === "checkUserBookingStatus") {
      result = checkUserBookingStatus(data.email);

    } else if (action === "getMentorProfile") {
      result = getMentorProfile(data.mentorName, data.email);

    } else if (action === "bookSlot") {
      result = bookSlot(data.mentor, data.email);  // ← note: "mentor" not "mentorName"

    }
    else if (action === "markBookingDone") {
       result = markBookingDone(data);
    }
     else {
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
  
  // Basic email format validation
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    return { valid: false, message: "Please enter a valid email address." };
  }
  
  // Check for VCU domain
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
      return { 
        success: false, 
        message: validation.message 
      };
    }
    
    var validEmail = validation.email;
    var verificationCode = generateVerificationCode();
    
    // Store verification code temporarily (expires in 10 minutes)
    var expirationTime = new Date().getTime() + (10 * 60 * 1000); // 10 minutes
    PropertiesService.getScriptProperties().setProperty(
      'verification_' + validEmail, 
      JSON.stringify({
        code: verificationCode,
        expires: expirationTime
      })
    );
    
    // Send verification email
    var emailSubject = "VCU Engineering Alternate Assignment Career Conversation - Verification Code";
    var emailBody = "Dear VCU Student,\n\nYour verification code for VCU Engineering Alternate Assignment Career Conversation is:\n\n🔐 VERIFICATION CODE: " + verificationCode + "\n\nThis code will expire in 10 minutes. Please enter this code on the website to access the mentor booking system.\n\nIf you did not request this code, please ignore this email.\n\nBest regards,\nVCU College of Engineering Career Services Team\n\n---\nThis is an automated message. Please do not reply to this email.";
    
    try {
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
      return { 
        success: false, 
        message: validation.message 
      };
    }
    
    var validEmail = validation.email;
    
    // Get stored verification data
    var verificationData = PropertiesService.getScriptProperties().getProperty('verification_' + validEmail);
    
    if (!verificationData) {
      return { 
        success: false, 
        message: "No verification code found. Please request a new code." 
      };
    }
    
    var parsedData = JSON.parse(verificationData);
    var currentTime = new Date().getTime();
    
    // Check if code has expired
    if (currentTime > parsedData.expires) {
      // Clean up expired code
      PropertiesService.getScriptProperties().deleteProperty('verification_' + validEmail);
      return { 
        success: false, 
        message: "Verification code has expired. Please request a new code." 
      };
    }
    
    // Check if code matches
    if (code.trim() !== parsedData.code) {
      return { 
        success: false, 
        message: "Invalid verification code. Please check your email and try again." 
      };
    }
    
    // Code is valid - clean up verification data and create session
    PropertiesService.getScriptProperties().deleteProperty('verification_' + validEmail);
    
    // Create authenticated session
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
    
    // Create index map using exact column names from Excel file
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
      signedUp: headers.indexOf("SignedUpStudents"),
      conversations: -1
    };
    
    // Find email column (try both options)
    if (indexes.email === -1) {
      indexes.email = indexes.preferredEmail;
    }
    
    // Find company column (try both options)
    if (indexes.company === -1) {
      indexes.company = indexes.companyAlt;
    }
    
    // Find conversations column
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i].trim();
      if (header.includes("How many conversations") && header.includes("would you be open to having")) {
        indexes.conversations = i;
        break;
      }
    }
    
    // Use conversations column if slots column doesn't exist
    if (indexes.slots === -1 && indexes.conversations !== -1) {
      indexes.slots = indexes.conversations;
    }
    
    // Add Signed-Up Students column if it doesn't exist
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
    
    // Use cached data if it's less than 30 seconds old and not forced refresh
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
    
    // Process mentor data more efficiently
    var mentors = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // removed signedup because signed-up has integers here, not strings
    //  var signedUp = row[indexes.signedUp] ? 
      //  row[indexes.signedUp].map(function(e) { return e.trim(); }).filter(function(e) { return e !== ""; }) : [];
      
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
        signedUpStudents: parseInt(row[indexes.signedUp]) || 0,
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
 
// Get mentors - OPTIMIZED VERSION
function getMentors(userEmail) {
  try {
    // Validate the email format
    if(userEmail){
    var validation = validateVCUEmail(userEmail);
    if (!validation.valid) {
      throw new Error("Invalid email: " + validation.message);
    }
    
    var studentEmail = validation.email;
    }
    // Use cached data for better performance
    var mentors = getCachedMentorData();
    
    // Add user-specific booking status
    return mentors.map(function(mentor) {
      return {
        name: mentor.name,
        areaOfFocus: mentor.areaOfFocus,
        industry: mentor.industry,
        company: mentor.company,
        major: mentor.major,
        email: mentor.email,
        availableSlots: mentor.availableSlots,
        signedUpStudents: mentor.signedUpStudents,
       // isBookedByCurrentUser: mentor.signedUpStudents.indexOf(studentEmail) !== -1
      };
    });
    
  } catch (error) {
    console.error("Error in getMentors:", error);
    throw new Error("Failed to load mentors: " + error.message);
  }
}

// Get detailed mentor information - OPTIMIZED VERSION
function getMentorProfile(mentorName, userEmail) {
 
  try {
    if(userEmail){
    // Validate the user email
    var validation = validateVCUEmail(userEmail);
    if (!validation.valid) {
      throw new Error("Invalid email: " + validation.message);
    }
    
    var studentEmail = validation.email;
    }
    // Use cached data
    var mentors = getCachedMentorData();
    
    // Find the specific mentor
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
    
    // Clean up LinkedIn URL
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
      //isBookedByCurrentUser: mentor.signedUpStudents.indexOf(studentEmail) !== -1
    };
    
  } catch (error) {
    console.error("Error in getMentorProfile:", error);
    throw new Error("Failed to load mentor profile: " + error.message);
  }
}

// Book a slot - OPTIMIZED VERSION
function bookSlot(mentorName, userEmail) {
  try {
    // Validate the user email
    var validation = validateVCUEmail(userEmail);
    if (!validation.valid) {
      return { 
        success: false, 
        message: "Invalid email: " + validation.message 
      };
    }

    
    var studentEmail = validation.email;
    
    var indexes = getColumnIndexes();
    
    // Get fresh data for booking (no cache to avoid race conditions)
    var data = sheet.getDataRange().getValues();

    // // Check if student has already booked ANY mentor
    // for (var i = 1; i < data.length; i++) {
    //   var signedUp = data[i][indexes.signedUp] || "";
    //   var signedUpList = signedUp.split(",").map(function(e) { return e.trim(); });
    //   if (signedUpList.indexOf(studentEmail) !== -1) {
    //     return { 
    //       success: false, 
    //       message: "You have already booked a session with another mentor. Each student can only book one mentor." 
    //     };
    //   }
    // }
    // Count how many mentors the student has booked
    //var studentBookingCount = 0;
    //for (var i = 1; i < data.length; i++) {
      //var signedUp = data[i][indexes.signedUp] || "";
      // @ts-ignore
      //var signedUpList = signedUp.map(function(e) { return e.trim(); });
      //if (signedUpList.indexOf(studentEmail) !== -1) {
      //  studentBookingCount++;
      //}
    //}
    //if (studentBookingCount >= 2) {
      //return { 
        //success: false, 
        //message: "You have already booked sessions with two mentors. Each student can book up to 2 mentors only." 
      //};
    //}


    // Find the selected mentor and book
    for (var i = 1; i < data.length; i++) {
      if (data[i][indexes.name] === mentorName) {
        var availableSlots = parseInt(data[i][indexes.slots]) || 0;
        
        if (availableSlots <= 0) {
          return { 
            success: false, 
            message: "This mentor has no available slots remaining." 
          };
        }
        
        // Update spreadsheet - decrease slots and add student
        sheet.getRange(i + 1, indexes.slots + 1).setValue(availableSlots - 1);
        
        var currentStudents = data[i][indexes.signedUp] || "";
        var updatedStudents = currentStudents + (currentStudents ? ", " : "") + studentEmail;
        sheet.getRange(i + 1, indexes.signedUp + 1).setValue(updatedStudents);
        
        // Clear cache after modification
        clearMentorDataCache();
        
        // Get mentor details for email (including LinkedIn)
        var mentorEmail = data[i][indexes.email];
        var areaOfFocus = data[i][indexes.areaOfFocus];
        var industry = data[i][indexes.industry];
        var company = data[i][indexes.company];
        
        // Get LinkedIn profile
        var linkedinUrl = data[i][indexes.linkedin] || "";
        // Clean up LinkedIn URL - add https if missing
        if (linkedinUrl && linkedinUrl.indexOf("http") !== 0) {
          linkedinUrl = "https://" + linkedinUrl;
        }
        
        // Send confirmation email
        var emailSubject = "VCU Engineering Alternate Assignment Career Conversation Mentor Confirmed - " + mentorName;
        var linkedinLine = linkedinUrl ? "\nLinkedIn Profile: " + linkedinUrl : "";
        var emailBody = `<div style="color: black;">
          <p>Dear <a href="mailto:${studentEmail}">${studentEmail}</a>,</p>

          <p><strong>==&gt; CONGRATULATIONS! &lt;==</strong><br>
          You have selected <strong>${mentorName}</strong> as your mentor for career conversation.</p>

          <h3>📋 MENTOR DETAILS</h3>
          <p>
            Name: ${mentorName}<br>
            Email: <a href="mailto:${mentorEmail}">${mentorEmail}</a><br>
            Area of Focus: ${areaOfFocus}<br>
            Industry: ${industry}<br>
            Company: ${company}<br>
            ${linkedinUrl ? 'LinkedIn Profile: <a href="' + linkedinUrl + '">' + linkedinUrl + '</a><br>' : ''}
          </p>

          <h3>▶ NEXT STEPS ◀</h3>
          <ol>
            <li><strong>REACH OUT TO ${mentorName} directly at: <a href="mailto:${mentorEmail}">${mentorEmail}</a></strong></li>
            <li>Schedule a convenient time for your 20-30 MINUTE VIDEO CALL</li>
            <li>The conversation should take place between MAR 2 AND APRIL 5</li>
            <li>Prepare thoughtful questions about their career path and industry</li>
          </ol>

          <h3>💡 SUGGESTED QUESTIONS FOR YOUR CONVERSATION</h3>

          <h4>TRENDS</h4>
          <ul>
            <li>What trends are most impacting your business/field right now?</li>
            <li>How has your business or field changed most since you started?</li>
            <li>How do you think your business or field will change most dramatically in the next several years?</li>
          </ul>

          <h4>INSIGHTS</h4>
          <ul>
            <li>What surprises you most about your job / employer?</li>
            <li>What's the best lesson you've learned on the job so far?</li>
            <li>What's been your best professional decision so far and why?</li>
            <li>If you had to attribute your success at your employer to one skill or trait, what would it be?</li>
          </ul>

          <h4>ADVICE</h4>
          <ul>
            <li>If you were me, what can I be doing right now to prepare myself for a career in this field?</li>
            <li>What do you know now that you wish you had known when you were in my position?</li>
            <li>If you were me, what would you be doing right now to maximize your chance of breaking into this industry/function?</li>
            <li>If you had just been hired into this role, what's the most important thing you would do in your first thirty days to ensure you got off to the fastest start possible?</li>
          </ul>

          <h4>RESOURCES</h4>
          <ul>
            <li>What resources do I need to look into next?</li>
            <li>What next steps would you recommend for someone in my situation?</li>
          </ul>

          <h4>ASSIGNMENT</h4>
          <ul>
            <li>Which project of yours do you feel has had the greatest impact?</li>
            <li>Has any particular type of project increased in popularity recently at your organization?</li>
            <li>Have you used interns or contractors in the past? If so, what sort of projects have they done?</li>
          </ul>

          <p>Feel free to use these questions as a starting point and adapt them based on your specific interests and the mentor's background.</p>

          <p>We're excited for you to connect with ${mentorName}!</p>

          <p>Best regards,<br>VCU COLLEGE OF ENGINEERING CAREER SERVICES TEAM</p>

          <hr>

          <p style="background-color: yellow; padding: 5px;">
            This is an automated message. Please do not reply to this email.
          </p></div>
        `;


        try {
          MailApp.sendEmail({
            to: studentEmail,
            subject: emailSubject,
            htmlBody: emailBody
          });
        } catch (emailError) {
          console.error("Email sending failed:", emailError);
          // Don't fail the booking if email fails
        }
        
        return { 
          success: true, 
          message: "Successfully booked with " + mentorName + "! Check your VCU email (" + studentEmail + ") for confirmation details." 
        };
      }
    }
    
    return { 
      success: false, 
      message: "Mentor not found or no longer available." 
    };
    
  } catch (error) {
    console.error("Error in bookSlot:", error);
    return { 
      success: false, 
      message: "Booking failed: " + error.message 
    };
  }
}

// Check if user has already booked a mentor - OPTIMIZED VERSION
function checkUserBookingStatus(userEmail) {
  try {
    var validation = validateVCUEmail(userEmail);
    if (!validation.valid) {
      return { hasBooked: false, error: "Invalid email" };
    }
    
    var studentEmail = validation.email;
    
    // Use cached data for checking booking status
    var mentors = getCachedMentorData();
    
    for (var i = 0; i < mentors.length; i++) {
      if (mentors[i].signedUpStudents.indexOf(studentEmail) !== -1) {
        return {
          hasBooked: true,
          mentorName: mentors[i].name
        };
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
    var mentorName   = data.mentor;
    var studentEmail = data.email;

    if (!mentorName || !studentEmail) {
      return { success: false, message: "Missing mentor or email" };
    }

    var indexes = getColumnIndexes(); // ← use same helper as bookSlot
    var values  = sheet.getDataRange().getValues();

    for (var i = 1; i < values.length; i++) {
      if (values[i][indexes.name] === mentorName) {

        var currentStudents = values[i][indexes.signedUp]
                              ? values[i][indexes.signedUp].toString()
                              : "";

        // Remove this student's email
        var updatedList = currentStudents
          .split(",")
          .map(function(s) { return s.trim(); })
          .filter(function(s) { return s !== studentEmail && s !== ""; })
          .join(", ");

        sheet.getRange(i + 1, indexes.signedUp + 1).setValue(updatedList);
        clearMentorDataCache();

        return { success: true, message: "Marked as done" };
      }
    }

    return { success: false, message: "Mentor not found: " + mentorName };

  } catch(e) {
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
      console.log("Column " + index + ": '" + header + "'");
    });
    
    return {
      totalColumns: headers.length,
      totalRows: data.length,
      headers: headers,
      indexes: indexes,
      success: true
    };
    
  } catch (error) {
    console.error("Debug error:", error);
    return { success: false, error: error.message };
  }
}
