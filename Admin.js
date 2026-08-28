// AdminCode.gs - All admin-related functions
// Admin password = vcu2026admin
// This file handles the admin dashboard functionality

// Get the same spreadsheet reference as the main code
var adminSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mentors");

// Simple admin authentication
function authenticateAdmin(password) {
  try {
    // Updated admin password
    const adminPassword = "vcu2026admin";
    
    if (password === adminPassword) {
      return {
        success: true,
        message: "Admin access granted"
      };
    } else {
      return {
        success: false,
        message: "Invalid admin password"
      };
    }
  } catch (error) {
    console.error("Admin auth error:", error);
    return {
      success: false,
      message: "Authentication failed"
    };
  }
}

// Get overview data for admin dashboard
function getAdminOverview() {
  try {
    var data = adminSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        totalMentors: 0,
        totalBookings: 0,
        availableSlots: 0,
        mentors: []
      };
    }
    
    var headers = data[0];
    console.log("Admin: Found headers:", headers);
    
    // Use the correct column mapping from your Excel file
    var nameIdx = headers.indexOf("First Name & Last Name");
    var areaOfFocusIdx = headers.indexOf("What is your area of focus?");
    var industryIdx = headers.indexOf("Industry you can share about.");
    
    // Company column mapping
    var companyIdx = headers.indexOf("Company");
    if (companyIdx === -1) {
      companyIdx = headers.indexOf("What company are you with or previously with?");
    }
    
    // Try flexible matching for company column
    if (companyIdx === -1) {
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i].trim();
        if (header === "What company are you with or previously with?") {
          companyIdx = i;
          break;
        }
      }
    }
    
    // Slots and signup columns
    var slotIdx = headers.indexOf("Available Slots");
    var signupIdx = headers.indexOf("Signed-Up Students");
    
    // Try alternative column names for slots
    if (slotIdx === -1) {
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i].trim();
        if (header.includes("How many conversations") && header.includes("would you be open to having")) {
          slotIdx = i;
          break;
        }
      }
    }
    
    // If Signed-Up Students column doesn't exist, add it
    if (signupIdx === -1) {
      console.log("Adding 'Signed-Up Students' column for admin...");
      var lastColumn = headers.length + 1;
      adminSheet.getRange(1, lastColumn).setValue("Signed-Up Students");
      signupIdx = lastColumn - 1;
      
      // Refresh data
      data = adminSheet.getDataRange().getValues();
      headers = data[0];
    }
    
    console.log("Admin column mapping:", {
      nameIdx: nameIdx,
      companyIdx: companyIdx,
      slotIdx: slotIdx,
      signupIdx: signupIdx
    });
    
    var mentors = [];
    var totalBookings = 0;
    var totalAvailableSlots = 0;
    
    // Process each mentor row
    for (var i = 1; i < data.length; i++) {
      var signedUpRaw = data[i][signupIdx] || "";
      var signedUp = signedUpRaw ? signedUpRaw.split(",").map(e => e.trim()).filter(e => e !== "") : [];
      var availableSlots = parseInt(data[i][slotIdx]) || 0;
      
      totalBookings += signedUp.length;
      totalAvailableSlots += availableSlots;
      
      mentors.push({
        name: data[i][nameIdx] || "",
        areaOfFocus: data[i][areaOfFocusIdx] || "",
        industry: data[i][industryIdx] || "",
        company: data[i][companyIdx] || "",
        availableSlots: availableSlots,
        signedUpStudents: signedUp,
        rowIndex: i + 1 // Store row index for updates later
      });
    }
    
    console.log(`Admin: Processed ${mentors.length} mentors, ${totalBookings} total bookings`);
    
    return {
      totalMentors: mentors.length,
      totalBookings: totalBookings,
      availableSlots: totalAvailableSlots,
      mentors: mentors
    };
    
  } catch (error) {
    console.error("Error in getAdminOverview:", error);
    throw new Error("Failed to load admin data: " + error.message);
  }
}

// Get detailed information about a specific mentor for admin management
function getAdminMentorDetails(mentorName) {
  try {
    var data = adminSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Get all column indexes using correct names from Excel file
    var nameIdx = headers.indexOf("First Name & Last Name");
    var areaOfFocusIdx = headers.indexOf("What is your area of focus?");
    var industryIdx = headers.indexOf("Industry you can share about.");
    var majorIdx = headers.indexOf("What was your major?");
    var additionalInfoIdx = headers.indexOf("Any other information about yourself that might be helpful to a student in determining whom to talk with? ");
    var linkedinIdx = headers.indexOf("What is your LinkedIn Profile?");
    
    // Company and email columns
    var companyIdx = headers.indexOf("Company");
    if (companyIdx === -1) {
      companyIdx = headers.indexOf("What company are you with or previously with?");
    }
    
    var emailIdx = headers.indexOf("Email Address");
    if (emailIdx === -1) {
      emailIdx = headers.indexOf("Preferred email address for students to use to contact you. This email will only be shared with the specific student(s) who sign up to talk with you. ");
    }
    
    var slotIdx = headers.indexOf("Available Slots");
    var signupIdx = headers.indexOf("Signed-Up Students");
    
    // Find the mentor
    for (var i = 1; i < data.length; i++) {
      if (data[i][nameIdx] === mentorName) {
        var signedUp = data[i][signupIdx] ? data[i][signupIdx].split(",").map(e => e.trim()).filter(e => e !== "") : [];
        
        return {
          name: data[i][nameIdx] || "",
          areaOfFocus: data[i][areaOfFocusIdx] || "",
          industry: data[i][industryIdx] || "",
          company: data[i][companyIdx] || "",
          major: data[i][majorIdx] || "",
          additionalInfo: data[i][additionalInfoIdx] || "",
          linkedinUrl: data[i][linkedinIdx] || "",
          email: data[i][emailIdx] || "",
          availableSlots: parseInt(data[i][slotIdx]) || 0,
          signedUpStudents: signedUp,
          rowIndex: i + 1
        };
      }
    }
    
    throw new Error("Mentor not found: " + mentorName);
    
  } catch (error) {
    console.error("Error in getAdminMentorDetails:", error);
    throw new Error("Failed to get mentor details: " + error.message);
  }
}

// Remove a student from a mentor (admin function)
function removeStudentFromMentor(mentorName, studentEmail) {
  try {
    var data = adminSheet.getDataRange().getValues();
    var headers = data[0];
    
    var nameIdx = headers.indexOf("First Name & Last Name");
    var slotIdx = headers.indexOf("Available Slots");
    var signupIdx = headers.indexOf("Signed-Up Students");
    
    // Find alternative slots column if needed
    if (slotIdx === -1) {
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i].trim();
        if (header.includes("How many conversations") && header.includes("would you be open to having")) {
          slotIdx = i;
          break;
        }
      }
    }
    
    // Find the mentor
    for (var i = 1; i < data.length; i++) {
      if (data[i][nameIdx] === mentorName) {
        var signedUpRaw = data[i][signupIdx] || "";
        var signedUpList = signedUpRaw.split(",").map(e => e.trim()).filter(e => e !== "");
        
        // Remove the student
        var updatedList = signedUpList.filter(email => email !== studentEmail);
        
        if (updatedList.length === signedUpList.length) {
          return {
            success: false,
            message: "Student not found in this mentor's list"
          };
        }
        
        // Update the spreadsheet
        var updatedStudentsString = updatedList.join(", ");
        adminSheet.getRange(i + 1, signupIdx + 1).setValue(updatedStudentsString);
        
        // Increase available slots by 1
        var currentSlots = parseInt(data[i][slotIdx]) || 0;
        adminSheet.getRange(i + 1, slotIdx + 1).setValue(currentSlots + 1);
        
        return {
          success: true,
          message: `Successfully removed ${studentEmail} from ${mentorName}`
        };
      }
    }
    
    return {
      success: false,
      message: "Mentor not found"
    };
    
  } catch (error) {
    console.error("Error in removeStudentFromMentor:", error);
    return {
      success: false,
      message: "Failed to remove student: " + error.message
    };
  }
}

// Add a student to a mentor manually (admin function)
function addStudentToMentor(mentorName, studentEmail) {
  try {
    // Validate email format
    var emailValidation = validateVCUEmail(studentEmail);
    if (!emailValidation.valid) {
      return {
        success: false,
        message: emailValidation.message
      };
    }
    
    var validEmail = emailValidation.email;
    var data = adminSheet.getDataRange().getValues();
    var headers = data[0];
    
    var nameIdx = headers.indexOf("First Name & Last Name");
    var slotIdx = headers.indexOf("Available Slots");
    var signupIdx = headers.indexOf("Signed-Up Students");
    
    // Find alternative slots column if needed
    if (slotIdx === -1) {
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i].trim();
        if (header.includes("How many conversations") && header.includes("would you be open to having")) {
          slotIdx = i;
          break;
        }
      }
    }
    
    // // Check if student is already booked with ANY mentor
    // for (var i = 1; i < data.length; i++) {
    //   var signedUp = data[i][signupIdx] || "";
    //   if (signedUp.split(",").map(e => e.trim()).includes(validEmail)) {
    //     return {
    //       success: false,
    //       message: `${validEmail} is already booked with ${data[i][nameIdx]}`
    //     };
    //   }
    // }
    // Count how many mentors the student has booked
    var studentBookingCount = 0;
    for (var i = 1; i < data.length; i++) {
      var signedUp = data[i][indexes.signedUp] || "";
      var signedUpList = signedUp.split(",").map(function(e) { return e.trim(); });
      if (signedUpList.indexOf(studentEmail) !== -1) {
        studentBookingCount++;
      }
    }
    if (studentBookingCount >= 2) {
      return { 
        success: false, 
        message: "You have already booked sessions with two mentors. Each student can book up to 2 mentors only." 
      };
    }

    
    // Find the specific mentor and add student
    for (var i = 1; i < data.length; i++) {
      if (data[i][nameIdx] === mentorName) {
        var availableSlots = parseInt(data[i][slotIdx]) || 0;
        
        if (availableSlots <= 0) {
          return {
            success: false,
            message: "This mentor has no available slots"
          };
        }
        
        // Add the student
        var currentStudents = data[i][signupIdx] || "";
        var updatedStudents = currentStudents + (currentStudents ? ", " : "") + validEmail;
        adminSheet.getRange(i + 1, signupIdx + 1).setValue(updatedStudents);
        
        // Decrease available slots
        adminSheet.getRange(i + 1, slotIdx + 1).setValue(availableSlots - 1);
        
        return {
          success: true,
          message: `Successfully added ${validEmail} to ${mentorName}`
        };
      }
    }
    
    return {
      success: false,
      message: "Mentor not found"
    };
    
  } catch (error) {
    console.error("Error in addStudentToMentor:", error);
    return {
      success: false,
      message: "Failed to add student: " + error.message
    };
  }
}

// Add slots to a mentor (admin function)
function addSlotsToMentor(mentorName, additionalSlots) {
  try {
    var data = adminSheet.getDataRange().getValues();
    var headers = data[0];
    
    var nameIdx = headers.indexOf("First Name & Last Name");
    var slotIdx = headers.indexOf("Available Slots");
    
    // Find alternative slots column if needed
    if (slotIdx === -1) {
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i].trim();
        if (header.includes("How many conversations") && header.includes("would you be open to having")) {
          slotIdx = i;
          break;
        }
      }
    }
    
    // Find the mentor and update slots
    for (var i = 1; i < data.length; i++) {
      if (data[i][nameIdx] === mentorName) {
        var currentSlots = parseInt(data[i][slotIdx]) || 0;
        var newSlots = currentSlots + additionalSlots;
        
        adminSheet.getRange(i + 1, slotIdx + 1).setValue(newSlots);
        
        return {
          success: true,
          message: `Added ${additionalSlots} slot(s) to ${mentorName}. New total: ${newSlots}`
        };
      }
    }
    
    return {
      success: false,
      message: "Mentor not found"
    };
    
  } catch (error) {
    console.error("Error in addSlotsToMentor:", error);
    return {
      success: false,
      message: "Failed to add slots: " + error.message
    };
  }
}

// Get all mentor data for export
function getAllMentorData() {
  try {
    var data = adminSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }
    
    var headers = data[0];
    
    // Map all relevant columns using correct names from Excel file
    var nameIdx = headers.indexOf("First Name & Last Name");
    var areaOfFocusIdx = headers.indexOf("What is your area of focus?");
    var industryIdx = headers.indexOf("Industry you can share about.");
    var majorIdx = headers.indexOf("What was your major?");
    var additionalInfoIdx = headers.indexOf("Any other information about yourself that might be helpful to a student in determining whom to talk with? ");
    var linkedinIdx = headers.indexOf("What is your LinkedIn Profile?");
    
    // Company and email columns
    var companyIdx = headers.indexOf("Company");
    if (companyIdx === -1) {
      companyIdx = headers.indexOf("What company are you with or previously with?");
    }
    
    var emailIdx = headers.indexOf("Email Address");
    if (emailIdx === -1) {
      emailIdx = headers.indexOf("Preferred email address for students to use to contact you. This email will only be shared with the specific student(s) who sign up to talk with you. ");
    }
    
    var slotIdx = headers.indexOf("Available Slots");
    var signupIdx = headers.indexOf("Signed-Up Students");
    
    // Try alternative column names for slots
    if (slotIdx === -1) {
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i].trim();
        if (header.includes("How many conversations") && header.includes("would you be open to having")) {
          slotIdx = i;
          break;
        }
      }
    }
    
    var mentors = [];
    
    // Process each mentor row
    for (var i = 1; i < data.length; i++) {
      var signedUpRaw = data[i][signupIdx] || "";
      var signedUp = signedUpRaw ? signedUpRaw.split(",").map(e => e.trim()).filter(e => e !== "") : [];
      
      mentors.push({
        name: data[i][nameIdx] || "",
        areaOfFocus: data[i][areaOfFocusIdx] || "",
        industry: data[i][industryIdx] || "",
        company: data[i][companyIdx] || "",
        major: data[i][majorIdx] || "",
        additionalInfo: data[i][additionalInfoIdx] || "",
        linkedinUrl: data[i][linkedinIdx] || "",
        email: data[i][emailIdx] || "",
        availableSlots: parseInt(data[i][slotIdx]) || 0,
        signedUpStudents: signedUp,
        studentCount: signedUp.length,
        rowIndex: i + 1
      });
    }
    
    return mentors;
    
  } catch (error) {
    console.error("Error in getAllMentorData:", error);
    throw new Error("Failed to get all mentor data: " + error.message);
  }
}

// Bulk update mentor slots (admin function)
function bulkUpdateSlots(updates) {
  try {
    var data = adminSheet.getDataRange().getValues();
    var headers = data[0];
    
    var nameIdx = headers.indexOf("First Name & Last Name");
    var slotIdx = headers.indexOf("Available Slots");
    
    // Find alternative slots column if needed
    if (slotIdx === -1) {
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i].trim();
        if (header.includes("How many conversations") && header.includes("would you be open to having")) {
          slotIdx = i;
          break;
        }
      }
    }
    
    var updateCount = 0;
    var errors = [];
    
    // Process each update
    updates.forEach(function(update) {
      var mentorName = update.mentorName;
      var newSlots = parseInt(update.slots);
      
      if (isNaN(newSlots) || newSlots < 0) {
        errors.push(`Invalid slot count for ${mentorName}: ${update.slots}`);
        return;
      }
      
      // Find and update the mentor
      for (var i = 1; i < data.length; i++) {
        if (data[i][nameIdx] === mentorName) {
          adminSheet.getRange(i + 1, slotIdx + 1).setValue(newSlots);
          updateCount++;
          break;
        }
      }
    });
    
    return {
      success: true,
      message: `Updated ${updateCount} mentors successfully`,
      updateCount: updateCount,
      errors: errors
    };
    
  } catch (error) {
    console.error("Error in bulkUpdateSlots:", error);
    return {
      success: false,
      message: "Failed to bulk update slots: " + error.message
    };
  }
}

// Get mentor statistics for admin dashboard
function getMentorStatistics() {
  try {
    var mentors = getAllMentorData();
    
    var stats = {
      totalMentors: mentors.length,
      totalSlots: 0,
      totalBookings: 0,
      availableSlots: 0,
      fullMentors: 0,
      emptyMentors: 0,
      industryBreakdown: {},
      companyBreakdown: {},
      averageBookingsPerMentor: 0,
      mostBookedMentor: { name: "", bookings: 0 },
      leastBookedMentor: { name: "", bookings: 999 }
    };
    
    mentors.forEach(function(mentor) {
      stats.totalSlots += mentor.availableSlots;
      stats.totalBookings += mentor.studentCount;
      stats.availableSlots += mentor.availableSlots;
      
      if (mentor.availableSlots === 0) {
        stats.fullMentors++;
      }
      
      if (mentor.studentCount === 0) {
        stats.emptyMentors++;
      }
      
      // Industry breakdown
      if (mentor.industry) {
        stats.industryBreakdown[mentor.industry] = (stats.industryBreakdown[mentor.industry] || 0) + 1;
      }
      
      // Company breakdown
      if (mentor.company) {
        stats.companyBreakdown[mentor.company] = (stats.companyBreakdown[mentor.company] || 0) + 1;
      }
      
      // Most/least booked
      if (mentor.studentCount > stats.mostBookedMentor.bookings) {
        stats.mostBookedMentor = { name: mentor.name, bookings: mentor.studentCount };
      }
      
      if (mentor.studentCount < stats.leastBookedMentor.bookings) {
        stats.leastBookedMentor = { name: mentor.name, bookings: mentor.studentCount };
      }
    });
    
    stats.averageBookingsPerMentor = mentors.length > 0 ? (stats.totalBookings / mentors.length).toFixed(1) : 0;
    
    return stats;
    
  } catch (error) {
    console.error("Error in getMentorStatistics:", error);
    throw new Error("Failed to get mentor statistics: " + error.message);
  }
}

// Search and filter mentors for admin
function searchMentors(query, filters) {
  try {
    var mentors = getAllMentorData();
    
    if (!query && !filters) {
      return mentors;
    }
    
    var filtered = mentors.filter(function(mentor) {
      var matchesQuery = true;
      var matchesFilters = true;
      
      // Text search
      if (query) {
        var searchText = query.toLowerCase();
        matchesQuery = 
          mentor.name.toLowerCase().includes(searchText) ||
          mentor.company.toLowerCase().includes(searchText) ||
          mentor.industry.toLowerCase().includes(searchText) ||
          mentor.areaOfFocus.toLowerCase().includes(searchText) ||
          mentor.major.toLowerCase().includes(searchText);
      }
      
      // Filters
      if (filters) {
        if (filters.industry && mentor.industry !== filters.industry) {
          matchesFilters = false;
        }
        
        if (filters.company && mentor.company !== filters.company) {
          matchesFilters = false;
        }
        
        if (filters.hasSlots !== undefined) {
          if (filters.hasSlots && mentor.availableSlots === 0) {
            matchesFilters = false;
          } else if (!filters.hasSlots && mentor.availableSlots > 0) {
            matchesFilters = false;
          }
        }
        
        if (filters.hasStudents !== undefined) {
          if (filters.hasStudents && mentor.studentCount === 0) {
            matchesFilters = false;
          } else if (!filters.hasStudents && mentor.studentCount > 0) {
            matchesFilters = false;
          }
        }
      }
      
      return matchesQuery && matchesFilters;
    });
    
    return filtered;
    
  } catch (error) {
    console.error("Error in searchMentors:", error);
    throw new Error("Failed to search mentors: " + error.message);
  }
}

// Reset all mentor data (admin emergency function)
function resetAllMentorData() {
  try {
    var data = adminSheet.getDataRange().getValues();
    var headers = data[0];
    
    var signupIdx = headers.indexOf("Signed-Up Students");
    var slotIdx = headers.indexOf("Available Slots");
    
    // Find alternative slots column if needed
    if (slotIdx === -1) {
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i].trim();
        if (header.includes("How many conversations") && header.includes("would you be open to having")) {
          slotIdx = i;
          break;
        }
      }
    }
    
    var resetCount = 0;
    
    // Clear all student bookings and reset slots to original values
    for (var i = 1; i < data.length; i++) {
      // Clear signed up students
      if (signupIdx !== -1) {
        adminSheet.getRange(i + 1, signupIdx + 1).setValue("");
      }
      
      // Reset available slots to original conversation count
      var originalSlots = parseInt(data[i][slotIdx]) || 0;
      if (slotIdx !== -1) {
        // This assumes the original slot count is stored somewhere
        // You might need to adjust this logic based on your needs
        adminSheet.getRange(i + 1, slotIdx + 1).setValue(originalSlots);
      }
      
      resetCount++;
    }
    
    return {
      success: true,
      message: `Reset ${resetCount} mentor records`,
      resetCount: resetCount
    };
    
  } catch (error) {
    console.error("Error in resetAllMentorData:", error);
    return {
      success: false,
      message: "Failed to reset mentor data: " + error.message
    };
  }
}function myFunction() {
  
}
