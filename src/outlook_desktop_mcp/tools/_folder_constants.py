"""Outlook OlDefaultFolders enum values and friendly-name mapping."""

OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_DRAFTS = 16
OL_FOLDER_CALENDAR = 9
OL_FOLDER_TASKS = 13
OL_FOLDER_JUNK = 23
OL_FOLDER_OUTBOX = 4

FOLDER_NAME_TO_ENUM = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT_MAIL,
    "sentmail": OL_FOLDER_SENT_MAIL,
    "deleted": OL_FOLDER_DELETED_ITEMS,
    "trash": OL_FOLDER_DELETED_ITEMS,
    "drafts": OL_FOLDER_DRAFTS,
    "calendar": OL_FOLDER_CALENDAR,
    "tasks": OL_FOLDER_TASKS,
    "junk": OL_FOLDER_JUNK,
    "spam": OL_FOLDER_JUNK,
    "outbox": OL_FOLDER_OUTBOX,
}

OL_MAIL_ITEM = 0
OL_APPOINTMENT_ITEM = 1

# OlBusyStatus
OL_BUSY_FREE = 0
OL_BUSY_TENTATIVE = 1
OL_BUSY_BUSY = 2
OL_BUSY_OUT_OF_OFFICE = 3
OL_BUSY_WORKING_ELSEWHERE = 4

# OlMeetingStatus
OL_NON_MEETING = 0
OL_MEETING = 1
OL_MEETING_RECEIVED = 3
OL_MEETING_CANCELED = 5

# OlMeetingResponse
OL_RESPONSE_TENTATIVE = 2
OL_RESPONSE_ACCEPTED = 3
OL_RESPONSE_DECLINED = 4

# OlRecipientType (for meetings)
OL_REQUIRED = 1
OL_OPTIONAL = 2
OL_RESOURCE = 3

BUSY_STATUS_NAMES = {
    0: "free", 1: "tentative", 2: "busy",
    3: "out_of_office", 4: "working_elsewhere",
}

MEETING_STATUS_NAMES = {
    0: "appointment", 1: "meeting", 3: "received", 5: "canceled",
}

RESPONSE_NAMES = {
    0: "none", 1: "organized", 2: "tentative",
    3: "accepted", 4: "declined", 5: "not_responded",
}
