const Imap = require("imap");
const { simpleParser } = require("mailparser");
const dotenv = require("dotenv");
const { GoogleGenerativeAI } = require("@google/generative-ai");
const moment = require("moment");
const { google } = require("googleapis");
dotenv.config();

const { default: fetch, Headers } = require("node-fetch");
globalThis.fetch = fetch;
globalThis.Headers = Headers;

// IMAP reconnection settings
const MAX_RECONNECT_ATTEMPTS = 5;
const RECONNECT_DELAY = 10000; // 10 seconds
const CHECK_INTERVAL = 10000; // 60 seconds
let imapReconnectAttempts = 0;

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
const model = genAI.getGenerativeModel({ model: "gemini-1.5-pro" });

// Set up Google Calendar auth
const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  process.env.GOOGLE_REDIRECT_URI
);

oauth2Client.setCredentials({
  refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
});

const calendar = google.calendar({ version: "v3", auth: oauth2Client });

// IMAP configuration
const imapConfig = {
  user: process.env.EMAIL,
  password: process.env.APP_PASSWORD,
  host: "imap.gmail.com",
  port: 993,
  tls: true,
  tlsOptions: { rejectUnauthorized: false },
  keepalive: true,
  authTimeout: 30000,
};

// Create IMAP connection
const imap = new Imap(imapConfig);

// Error handling with reconnection
function handleImapError(error) {
  console.error(`IMAP Error: ${error}`);

  if (imapReconnectAttempts < MAX_RECONNECT_ATTEMPTS) {
    imapReconnectAttempts++;
    console.log(
      `Connection lost. Attempting to reconnect (attempt ${imapReconnectAttempts} of ${MAX_RECONNECT_ATTEMPTS})...`
    );

    try {
      if (imap && imap.state !== "disconnected") {
        imap.end();
      }
    } catch (e) {
      console.log(`Error ending existing connection: ${e.message}`);
    }

    setTimeout(() => {
      console.log(`Reconnecting to IMAP server...`);
      imap.connect();
    }, RECONNECT_DELAY);
  } else {
    console.log(
      `Maximum reconnection attempts (${MAX_RECONNECT_ATTEMPTS}) reached. Giving up.`
    );
    setTimeout(() => {
      imapReconnectAttempts = 0;
    }, 60000); // Reset after 1 minute
  }
}

// Event handlers for IMAP connection
imap.once("error", (err) => {
  console.error("IMAP Error:", err);
  if (err.source === "timeout") {
    console.log("Reconnecting in 10 seconds...");
    setTimeout(() => imap.connect(), 10000);
  } else {
    handleImapError(err);
  }
});

imap.once("end", () => {
  console.log("IMAP Connection closed. Reconnecting...");
  setTimeout(() => imap.connect(), 5000);
});

imap.once("ready", function () {
  openInbox((err, box) => {
    if (err) {
      console.error("Error opening inbox:", err);
      return;
    }

    console.log(`Listening for new emails in ${box.name}...`);

    // Set up periodic checking
    setInterval(() => {
      // console.log('Periodically checking for new emails...');
      checkEmails();
    }, CHECK_INTERVAL);

    // Listen for new emails
    imap.on("mail", function () {
      console.log("New email detected...");
      checkEmails();
    });
  });
});

// Open the inbox/label
function openInbox(cb) {
  const labelName = process.env.LABEL_LISTENING; // Change this to your desired label
  imap.openBox(labelName, false, cb);
}

// Check for new emails
function checkEmails() {
  imap.search(["UNSEEN"], (err, results) => {
    if (err) {
      console.error("Error searching for unread emails:", err);
      return;
    }

    if (!results.length) {
      // console.log('No unread emails found');
      return;
    }

    console.log(`Found ${results.length} unread email(s)`);

    const fetch = imap.fetch(results, {
      bodies: [
        "HEADER.FIELDS (SUBJECT FROM DATE REFERENCES MESSAGE-ID X-GM-THRID)",
        "TEXT",
      ],
      markSeen: true,
      struct: true,
    });

    // Track emails by sequence number
    const emailsData = new Map();

    fetch.on("message", (msg, seqno) => {
      console.log(`Processing message #${seqno}`);

      // Initialize email data structure
      const emailData = {
        subject: "",
        fromEmail: "",
        body: "",
        date: "",
        threadId: "",
        partsProcessed: 0,
        totalParts: 2, // We expect header and body parts
      };

      // Store in our tracking map
      emailsData.set(seqno, emailData);

      // Extract threadId from attributes
      msg.on("attributes", (attrs) => {
        if (attrs["x-gm-thrid"]) {
          emailData.threadId = attrs["x-gm-thrid"].toString();
        }
      });

      // Process each part of the message
      msg.on("body", (stream, info) => {
        let buffer = "";

        stream.on("data", (chunk) => {
          buffer += chunk.toString("utf8");
        });

        stream.on("end", () => {
          // Process header fields
          if (
            info.which ===
            "HEADER.FIELDS (SUBJECT FROM DATE REFERENCES MESSAGE-ID X-GM-THRID)"
          ) {
            const header = Imap.parseHeader(buffer);
            emailData.subject = header.subject?.[0] || "";
            emailData.date = header.date?.[0] || "";

            // Use message-id as fallback for threadId
            if (!emailData.threadId && header["message-id"]?.[0]) {
              emailData.threadId = header["message-id"][0].replace(/[<>]/g, "");
            }

            // Extract email from the from field
            const fromField = header.from?.[0] || "";
            emailData.fromEmail = fromField.match(/<(.+)>/)?.[1] || fromField;

            emailData.partsProcessed++;
          }
          // Process email body
          else if (info.which === "TEXT") {
            // Use simpleParser to extract text content
            simpleParser(buffer, (err, parsed) => {
              if (err) {
                console.error(`Error parsing email body: ${err.message}`);
                emailData.body = buffer; // Fallback to raw buffer
              } else {
                emailData.body = parsed.text || buffer;
              }

              emailData.partsProcessed++;
            });
          }
        });
      });

      msg.once("end", () => {
        console.log(`Finished fetching all parts for message #${seqno}`);
      });
    });

    fetch.once("error", (err) => {
      console.error("Error fetching emails:", err);
    });

    fetch.once("end", () => {
      console.log("Done fetching all messages");

      // Process completed emails after a short delay
      // This ensures all async parsers have completed
      setTimeout(() => {
        // Filter for emails with complete data
        const completeEmails = Array.from(emailsData.values()).filter(
          (email) => {
            const isComplete =
              email.partsProcessed === email.totalParts &&
              email.threadId &&
              email.fromEmail &&
              email.subject !== undefined &&
              email.body;

            if (!isComplete) {
              console.log(
                `Skipping incomplete email: ${JSON.stringify({
                  threadId: email.threadId,
                  fromEmail: email.fromEmail,
                  partsProcessed: email.partsProcessed,
                  totalParts: email.totalParts,
                })}`
              );
            }

            return isComplete;
          }
        );

        // Add debugging here - after the delay and processing
        // console.log("Complete email data:", JSON.stringify(completeEmails, null, 2));

        console.log(
          `Processing ${completeEmails.length} complete emails out of ${emailsData.size} total`
        );

        // Process each complete email
        completeEmails.forEach(async (email) => {
          try {
            // First check if this email thread already has an event
            const existingEvent = await findExistingEvent(email.threadId);

            if (existingEvent) {
              // If there's an existing event, check for cancellations or updates
              console.log(
                `Found existing event for thread ID ${email.threadId}, checking for updates/cancellation`
              );
              await checkEventStatus(email, existingEvent);
            } else {
              // If there's no existing event, check if this is a new event
              console.log(
                `No existing event for thread ID ${email.threadId}, checking if this is a new event`
              );
              await processEvent(email);
            }
          } catch (error) {
            console.error(
              `Error processing email (threadId: ${email.threadId}): ${error.message}`
            );
          }
        });
      }, 2000);
    });
  });
}

// Add this processEvent function definition before it's called
async function processEvent(emailData) {
  // Add a small delay before each API call to avoid rate limiting
  await sleep(API_DELAY);

  const prompt = `You are given the body of an email sent out to a college student of Ashoka University. 
                        Your task is to identify whether or not the email is an event email. 
                        An event is defined as the following - 
                        An event is something that happens or is regarded as happening; an occurrence, especially one of some importance.

                        Some key aspects of what constitutes an event are Change, Specificity, Significance.
                        An event always has a date, time (compulsorily) and a venue (optionally, venue could be 'tbd' or unspecified) associated with it. Moreover, the language
                        in the email corresponds to a that of an event clearly. Note that 'deadlines' are not the same as timings of an events. Look out
                        for differences in date-time mentioned in deadlines v/s for events.
                        Once you have identified it, respond with a tuple. The first item in the tuple is a boolean True or False ONLY. Do not even change the case of the words.
                        If your answer is false, the second item is an empty object. 
                        If you answer is true, the second item is a JSON object containing specifications of the event.
                        Specifications of the event include: 
                        1. Name of the event
                        2. Organising Body (If the event is a collaboration between 2 or more bodies, mention all. Also, note that Ashoka University is my university, and it will be mentioned in every event - it should not be counted as an organising body)
                        3. Date, Time, Venue (If this is a word like 'Today' or 'Tomorrow', then use look at the timestamp of the email mentioning
                        this day to calculate the date of the event from the day and replace appropriately. Words should not appear in the JSON object, only proper dates.)
                        4. A concise descriptive summary of the event
                        Only extract names of organising bodies from the 'From' part provided to you. No other part of the email should be consulted for this. 
                        Moreover the descriptive summary should not involve you elaborating on any terms mentioned in the email. Simply summarise the email content. Do not elaborate or use your knowledge to explain the event at all. 
                        Moreover, when figuring out the venue, look out for the entire venue. For example if the event is 'in front of the mess' then the venue is 'in front of the mess', not just 'mess'. 
                        Similarly, if the event is taking place 'in the mess lawns', the venue is 'the mess lawns', not 'lawns'. Be liberal in selecting the venue, it can be a phrase too, not just a location. Also, the time should be in the 12 hour H:MM format with AM/PM.
                        Example of a valid JSON object is: 
                        {
                            "Name of the event": "AI and Ethics Symposium",
                            "Organising Body": ["Computer Science Department", "Centre for AI Policy"],
                            "Date, Time, Venue": {
                                "Date": "2025-03-15",
                                "Time": "10:00 AM - 4:00 PM (Note, if only start time is provided, add 1 hour to it for the end time)",
                                "Venue": "AC-02-LR-011"
                            },
                            "Descriptive Summary": "The AI and Ethics Symposium brings together leading experts, scholars, and students to explore the ethical implications of artificial intelligence. The event includes keynote speeches and poster presentations."
                        }
                        While determining "Time", make sure you NEVER enter any words except time in 12-hour format. Do not use words like "onwards". Only numbers separated by ':' and followed by 'AM' or 'PM' are permissible.
                        Make sure you appropriately close the braces in the JSON object and follow all specifications of how a JSON object should be.
                        Your answer should be ONLY the tuple. No other surrounding words or phrases. Pick up the date of event based on the date, year, month attached here: ${emailData.date}. Give me the time output in the h:mm AM/PM format.
                        The tuple must NOT be enclosed in brackets and there must be no other surrounding characters or words from you. Just write True/False followed by a comma and then the object. No other characters.
                        The email is provided to you below: 
                        Subject: ${emailData.subject} From: ${emailData.fromEmail}
                    Body: ${emailData.body}.`;

  try {
    // console.log(`Processing email with subject: "${emailData.subject}" from: ${emailData.fromEmail}`);

    const result = await model.generateContent(prompt);
    const llmResponse = result.response.text();
    // console.log(`LLM response received: ${llmResponse}`);

    // Parse the tuple response
    const match = llmResponse.match(/(True|False),\s*({.*}|\{\})/s);

    if (match) {
      const isEvent = match[1] === "True";
      let eventObject = match[2];

      try {
        if (typeof eventObject === "string") {
          eventObject = JSON.parse(eventObject);
        }

        // console.log(`Email is${isEvent ? '' : ' not'} an event`);

        if (isEvent) {
          // No need to check for existing events again - we already did that
          // Just create the event
          await sendToGoogleCalendar(
            eventObject,
            emailData.threadId,
            emailData.fromEmail
          );

          return {
            isEvent,
            eventData: eventObject,
            threadId: emailData.threadId,
          };
        } else {
          console.log("Not an event, skipping further processing");
          return {
            isEvent: false,
            eventData: null,
            threadId: emailData.threadId,
          };
        }
      } catch (error) {
        console.log(`Error parsing event object: ${error}`);
        return {
          isEvent: false,
          eventData: null,
          threadId: emailData.threadId,
          error: `Failed to parse event data: ${error.message}`,
        };
      }
    } else {
      console.log("Could not parse LLM response as a tuple");
      return {
        isEvent: false,
        eventData: null,
        threadId: emailData.threadId,
        error: "Invalid LLM response format",
      };
    }
  } catch (error) {
    console.log(`Error getting response from AI model: ${error}`);
    return {
      isEvent: false,
      eventData: null,
      threadId: emailData.threadId,
      error: `AI model error: ${error.message}`,
    };
  }
}

// Utility Functions
// Delay between API calls to avoid rate limiting
const API_DELAY = 1000; // 1 second delay between API calls

// Helper function to add delay
function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Check if an event with the given threadId already exists
async function findExistingEvent(threadId) {
  try {
    const response = await calendar.events.list({
      calendarId: process.env.GOOGLE_CALENDAR_ID,
      privateExtendedProperty: `threadId=${threadId}`,
    });

    if (response.data.items && response.data.items.length > 0) {
      return response.data.items[0]; // Return the first matching event
    }
    return null;
  } catch (error) {
    console.log(`Error finding existing event: ${error}`);
    return null;
  }
}

async function sendToGoogleCalendar(eventObject, threadId, fromEmail) {
  try {
    console.log("Event object received:", JSON.stringify(eventObject, null, 2));

    // Get date and time values
    const eventDate = eventObject["Date, Time, Venue"].Date;
    const eventTimeStr = eventObject["Date, Time, Venue"].Time;

    // Split the time range
    const timeParts = eventTimeStr.split(" - ");
    const startTimeStr = timeParts[0];
    const endTimeStr = timeParts.length > 1 ? timeParts[1] : null;

    // Parse date using moment
    const parsedDate = moment(eventDate);
    if (!parsedDate.isValid()) {
      console.error(`Invalid date: ${eventDate}, using current date`);
      parsedDate = moment();
    }

    // Parse start time correctly - DON'T concatenate strings
    // Use the moment object's methods instead
    const startTime = moment(startTimeStr, ["h:mm A", "h:mm a"]);
    if (!startTime.isValid()) {
      console.error(`Invalid start time: ${startTimeStr}, using current time`);
      startTime = moment();
    }

    // Create a proper datetime by setting hour and minute on the date
    const startDateTime = moment(parsedDate).set({
      hour: startTime.hour(),
      minute: startTime.minute(),
      second: 0,
    });

    // Handle end time similarly
    let endDateTime;
    if (endTimeStr) {
      const endTime = moment(endTimeStr, ["h:mm A", "h:mm a"]);
      if (endTime.isValid()) {
        endDateTime = moment(parsedDate).set({
          hour: endTime.hour(),
          minute: endTime.minute(),
          second: 0,
        });
      } else {
        console.error(
          `Invalid end time: ${endTimeStr}, using start time + 1 hour`
        );
        endDateTime = moment(startDateTime).add(1, "hour");
      }
    } else {
      console.log(`No end time provided, using start time + 1 hour`);
      endDateTime = moment(startDateTime).add(1, "hour");
    }

    // console.log(`Final start time: ${startDateTime.format('YYYY-MM-DD hh:mm A')} (${startDateTime.toISOString()})`);
    // console.log(`Final end time: ${endDateTime.format('YYYY-MM-DD hh:mm A')} (${endDateTime.toISOString()})`);

    // Build the event object that would be sent to Google Calendar
    const calendarEvent = {
      summary: eventObject["Name of the event"] || "Untitled Event",
      location: eventObject["Date, Time, Venue"].Venue || "",
      description:
        eventObject["Descriptive Summary"] || "No description available",
      start: {
        dateTime: startDateTime.toISOString(),
        timeZone: "Asia/Kolkata",
      },
      end: {
        dateTime: endDateTime.toISOString(),
        timeZone: "Asia/Kolkata",
      },
      attendees: fromEmail ? [{ email: fromEmail }] : [],
      extendedProperties: {
        private: {
          emailSource: fromEmail || "",
          threadId: threadId,
        },
        shared: {
          emailSource: fromEmail || "",
          threadId: threadId,
        },
      },
    };

    console.log(
      "Calendar event object:",
      JSON.stringify(calendarEvent, null, 2)
    );

    // Actually send to Google Calendar
    try {
      console.log(`Creating event in Google Calendar...`);
      const response = await calendar.events.insert({
        calendarId: process.env.GOOGLE_CALENDAR_ID,
        resource: calendarEvent,
      });

      console.log(`Event created successfully: ${response.data.htmlLink}`);
      return {
        status: "success",
        message: "Event created successfully",
        link: response.data.htmlLink,
        eventId: response.data.id,
      };
    } catch (apiError) {
      console.error(`Error calling Google Calendar API: ${apiError.message}`);

      if (apiError.response) {
        console.error(`Response status: ${apiError.response.status}`);
        console.error(`Response data:`, apiError.response.data);
      }

      return {
        status: "error",
        message: `Google Calendar API error: ${apiError.message}`,
      };
    }
  } catch (e) {
    console.log(`Error in sendToGoogleCalendar: ${e}`);
    console.log(e.stack); // Show the full stack trace
    return { status: "error", message: e.message };
  }
}

async function checkEventStatus(emailData, existingEvent) {
  // Add a small delay before each API call to avoid rate limiting
  await sleep(API_DELAY);

  const prompt = `Analyze the following email to determine:
                 1. If it's cancelling an event mentioned in a previous email
                 2. If it contains updates to an existing event

                 First, determine if the email is cancelling an event.
                 Look for keywords like (not exhaustively):
                 - "cancelled"
                 - "called off"
                 - "will not take place"
                 - "is cancelled"
                 - "we regret to inform"
                 - "unfortunately" combined with cancellation context
                 
                 Then, if the email is NOT cancelling an event, check if it contains updates to:
                 - Date 
                 - Time
                 - Venue
                 - Any other important event details
                 
                 Return your response as a JSON object in the following format:
                 {
                   "isCancelled": true/false,
                   "updates": {
                     // If isCancelled is true, this can be an empty object
                     // If isCancelled is false and there are no updates, this can be an empty object
                     // If isCancelled is false and there are updates, include the updated event details here:
                     "Name of the event": "Updated event name",
                     "Organising Body": ["Updated Organising Body"],
                     "Date, Time, Venue": {
                       "Date": "Updated date (YYYY-MM-DD)",
                       "Time": "Updated time in 12-hour format (e.g., '10:00 AM - 4:00 PM'). If only start time is provided, use the same duration as the original event object.",
                       "Venue": "Updated venue"
                     }
                   }
                 }
                 
                 Provide ONLY this JSON object in your response, with no surrounding text or phrases. Don't include \` or any other such characters, and you need not mention that this is a json.
                 
                 Email Subject: ${emailData.subject}
                 Email Body: ${emailData.body}
                 
                 Existing event details:
                 ${JSON.stringify(existingEvent, null, 2)}`;

  try {
    const result = await model.generateContent(prompt);
    const response = result.response.text().trim();

    try {
      // Parse the LLM response
      const parsedResponse = JSON.parse(response);

      console.log(
        `Event status check: isCancelled=${parsedResponse.isCancelled}`
      );

      if (parsedResponse.isCancelled) {
        console.log("Event cancellation detected");
        await removeCancelledEvent(emailData.threadId);
      } else if (Object.keys(parsedResponse.updates).length > 0) {
        console.log(
          "Event updates detected:",
          JSON.stringify(parsedResponse.updates, null, 2)
        );
        await applyUpdatesToEvent(existingEvent, parsedResponse.updates);
      } else {
        console.log("No cancellation or updates detected");
      }

      return parsedResponse;
    } catch (parseError) {
      console.error(`Error parsing LLM response: ${parseError.message}`);
      console.error(`Raw response: ${response}`);
      return { isCancelled: false, updates: {} };
    }
  } catch (error) {
    console.error(`Error checking event status: ${error.message}`);
    return { isCancelled: false, updates: {} };
  }
}

// Skeleton for removeCancelledEvent function
async function removeCancelledEvent(threadId) {
  try {
    console.log(`Removing cancelled event with threadId: ${threadId}`);
    const existingEvent = await findExistingEvent(threadId);

    if (existingEvent) {
      const response = await calendar.events.delete({
        calendarId: process.env.GOOGLE_CALENDAR_ID,
        eventId: existingEvent.id,
      });

      if (response.status === 204) {
        // Success status for deletion
        console.log(
          `Successfully deleted cancelled event with threadId: ${threadId}`
        );
        return true;
      } else {
        console.log(
          `Unexpected response status when deleting event: ${response.status}`
        );
        return false;
      }
    }

    console.log(`No existing event found with threadId: ${threadId}`);
    return false;
  } catch (error) {
    if (error.code === 404) {
      console.log(`Event not found (may have been already deleted)`);
    } else {
      console.log(`Error deleting cancelled event: ${error.message}`);
    }
    return false;
  }
}

// Skeleton for applyUpdatesToEvent function
async function applyUpdatesToEvent(existingEvent, updates) {
  console.log(`Applying updates to event ID: ${existingEvent.id}`);

  try {
    // Create a copy of the existing event to modify
    const updatedEvent = { ...existingEvent };
    let changesDetected = false;

    // Update name if provided
    if (updates["Name of the event"]) {
      updatedEvent.summary = updates["Name of the event"];
      console.log(`Updating event name to: ${updates["Name of the event"]}`);
      changesDetected = true;
    }

    // Update location/venue if provided
    if (updates["Date, Time, Venue"] && updates["Date, Time, Venue"].Venue) {
      // Skip "Not Mentioned" as it's not a real update
      if (updates["Date, Time, Venue"].Venue !== "Not Mentioned") {
        updatedEvent.location = updates["Date, Time, Venue"].Venue;
        console.log(`Updating venue to: ${updates["Date, Time, Venue"].Venue}`);
        changesDetected = true;
      }
    }

    // Process date and time updates
    if (updates["Date, Time, Venue"]) {
      // Get existing event times
      const startDateTime = moment(existingEvent.start.dateTime);
      const endDateTime = moment(existingEvent.end.dateTime);

      // Calculate the original duration in minutes
      const duration = endDateTime.diff(startDateTime, "minutes");
      console.log(`Original event duration: ${duration} minutes`);

      // Create new date/time objects that we'll modify if updates exist
      let newStartDateTime = moment(startDateTime);
      let newEndDateTime = moment(endDateTime);
      let dateUpdated = false;
      let timeUpdated = false;

      // Update date if provided
      if (updates["Date, Time, Venue"].Date) {
        let parsedDate = moment(updates["Date, Time, Venue"].Date);

        if (parsedDate.isValid()) {
          console.log(
            `Updating date from ${startDateTime.format(
              "YYYY-MM-DD"
            )} to ${parsedDate.format("YYYY-MM-DD")}`
          );

          // Update only the date components while preserving time
          newStartDateTime = moment(newStartDateTime).set({
            year: parsedDate.year(),
            month: parsedDate.month(),
            date: parsedDate.date(),
          });

          newEndDateTime = moment(newEndDateTime).set({
            year: parsedDate.year(),
            month: parsedDate.month(),
            date: parsedDate.date(),
          });

          dateUpdated = true;
          changesDetected = true;
        } else {
          console.error(
            `Invalid date format received: ${updates["Date, Time, Venue"].Date}`
          );
        }
      }

      // Update time if provided
      if (updates["Date, Time, Venue"].Time) {
        const timeParts = updates["Date, Time, Venue"].Time.split(" - ");
        const startTimeStr = timeParts[0].trim();

        // Parse the start time
        let startTime = moment(startTimeStr, [
          "h:mm A",
          "h:mm a",
          "hh:mm A",
          "hh:mm a",
        ]);

        if (startTime.isValid()) {
          console.log(`Updating start time to: ${startTimeStr}`);

          // Update only the time components
          newStartDateTime = moment(newStartDateTime).set({
            hour: startTime.hour(),
            minute: startTime.minute(),
            second: 0,
          });

          // If we have an end time, use it; otherwise, maintain the original duration
          if (timeParts.length > 1) {
            const endTimeStr = timeParts[1].trim();
            let endTime = moment(endTimeStr, [
              "h:mm A",
              "h:mm a",
              "hh:mm A",
              "hh:mm a",
            ]);

            if (endTime.isValid()) {
              console.log(`Updating end time to: ${endTimeStr}`);

              newEndDateTime = moment(newEndDateTime).set({
                hour: endTime.hour(),
                minute: endTime.minute(),
                second: 0,
              });
            } else {
              console.error(
                `Invalid end time format: ${endTimeStr}, maintaining original duration`
              );
              newEndDateTime = moment(newStartDateTime).add(
                duration,
                "minutes"
              );
            }
          } else {
            console.log(
              `No end time provided, maintaining original duration of ${duration} minutes`
            );
            newEndDateTime = moment(newStartDateTime).add(duration, "minutes");
          }

          timeUpdated = true;
          changesDetected = true;
        } else {
          console.error(`Invalid start time format: ${startTimeStr}`);
        }
      }

      // Apply the date/time updates
      if (dateUpdated || timeUpdated) {
        updatedEvent.start.dateTime = newStartDateTime.toISOString();
        updatedEvent.end.dateTime = newEndDateTime.toISOString();
        console.log(
          `Updated event time: ${newStartDateTime.format(
            "YYYY-MM-DD hh:mm A"
          )} to ${newEndDateTime.format("YYYY-MM-DD hh:mm A")}`
        );
      }
    }

    // Only update if changes were detected
    if (changesDetected) {
      console.log(`Changes detected, sending updated event to Google Calendar`);

      try {
        const response = await calendar.events.update({
          calendarId: process.env.GOOGLE_CALENDAR_ID,
          eventId: existingEvent.id,
          resource: updatedEvent,
        });

        console.log(`Event updated successfully: ${response.data.htmlLink}`);
        return {
          status: "success",
          message: "Event updated successfully",
          link: response.data.htmlLink,
        };
      } catch (apiError) {
        console.error(
          `Error updating event in Google Calendar: ${apiError.message}`
        );

        if (apiError.response) {
          console.error(`Response status: ${apiError.response.status}`);
          console.error(`Response data:`, apiError.response.data);
        }

        return {
          status: "error",
          message: `Google Calendar API error: ${apiError.message}`,
        };
      }
    } else {
      console.log(`No changes detected in the event, skipping update`);
      return {
        status: "no_change",
        message: "No changes detected in the event",
      };
    }
  } catch (error) {
    console.error(`Unexpected error in applyUpdatesToEvent: ${error.message}`);
    console.error(error.stack);
    return {
      status: "error",
      message: `Error applying updates: ${error.message}`,
    };
  }
}

imap.connect();
