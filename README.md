# outlook-ole-ruby

* Installation (Windows)

1) Install Chocolatey (https://chocolatey.org/install)
2) > choco install ruby
3) rename main_example.rb to main.rb
3) edit main.rb - create all the required actions

* Automatic sorting and cleanup

DELETE_RULES = [
  Rule.new(...),
  Rule.new(...),
  Rule.new(...),
  ...
]

is just an array of rules. Another one is MOVE_RULES. Each email either matches the rule or not. If it does, it will be deleted or moved by

outlook.autoclean_folder('Ivan@corp.com/Inbox', DELETE_RULES, MOVE_RULES)

the 1st parameter ('Ivan@corp.com/Inbox') is the source path (datastore/folder)
the 2nd parameter is an array of delete rules
the 3rd parameter is an array of move rules (optional)

Rules can have any of:
body: /body regexp/
subject: /subject regexp/
sender: 'sender@something.com'
sender_like: /sender regexp/
older_than_days: 10
type: any of :is_appointment?, :is_contact?, :is_meeting_request?, :is_meeting_cancelled?, :is_non_delivery_report?, :not_email?

Example rules:
Rule.new(body: /boring/) will matches all emails with the word "boring" in text
Rule.new(body: /boring/, subject: /Survey/) does the same, but only if the subject contains "Survey" word
Rule.new(sender: 'vasya.pupkin@gmail.com', older_than_days: 30) - Vasya's emails received a month+ ago
Rule.new(type: :is_non_delivery_report?, older_than_days: 1) - "email cannot be delivered" reports >1 day old

When you make a rule for MOVE_RULES array, also add move_to: 'new destination' (datastore/folder):
Rule.new(sender: 'do-not-reply@client.com', subject: /\[Development\]/, move_to: 'Ivan@corp.com/System Test')

* Automatic archiving

outlook.archive_mail('Ivan.Kokorev@excelian.com/Inbox', 'Title', 'c:/Users/vanya/Documents/Outlook', divide_by: :year, rules: ARCHIVE_RULES)

Downloads emails from the folder (1st parameter) to new PSTs named "Title 2019 08" (for divide_by: :month) or "Title 2019" (for divide_by: :year, which is default). Add an optional parameter rules: some_array_of_rules in case you don't want all emails to be archived.

* Running

Open cmd at the project folder and run "ruby main.rb". You can also add a task to Task Scheduler