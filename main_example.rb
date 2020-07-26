require_relative 'rule'
require_relative 'email'
require_relative 'outlook'

DELETE_RULES = [
  Rule.new(subject: /Some Words From Subject/, older_than_days: 3),
  Rule.new(sender: 'do-not-reply@client.com', older_than_days: 7),
  Rule.new(body: /mailbox is almost full/),
  Rule.new(type: :is_meeting_cancelled?, older_than_days: 7),
  Rule.new(type: :is_non_delivery_report?),
]

MOVE_RULES = [
  Rule.new(sender: 'do-not-reply@client.com', subject: /\[Development\]/, move_to: 'Ivan@corp.com/System Test'),
]

ARCHIVE_RULES = [
  Rule.new(sender_like: /@client.com/, older_than_days: 30)
  Rule.new(sender: 'masha@corp.com', older_than_days: 30)
]

outlook = Outlook.new
outlook.autoclean_folder('Ivan@corp.com/Inbox', DELETE_RULES, MOVE_RULES)
outlook.autoclean_folder('Ivan@corp.com/Proton Test', [Rule.new(older_than_days: 3)])
outlook.autoclean_folder('Ivan@corp.com/Proton Staging', [Rule.new(older_than_days: 3)])
outlook.autoclean_folder('Ivan@corp.com/Proton Prod', [Rule.new(older_than_days: 7)])
outlook.archive_mail('Ivan@corp.com/Inbox', 'Corp', 'c:/Users/vanya/Documents/Corp', divide_by: :year, rules: ARCHIVE_RULES)