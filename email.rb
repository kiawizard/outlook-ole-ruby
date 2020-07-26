class Email
  PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
  def initialize(email)
    @subject = email.subject
    @body = email.body
    @created_at = email.creationtime
    @age_days = ((Time.now - @created_at)/3600/24).floor
    @raw = email
    @type = email.messageclass
    @sender = get_sender(email)
  end

  def type
    @type
  end

  def not_email?
    is_appointment? || is_contact? || is_meeting_request? || is_meeting_cancelled? || is_non_delivery_report?
  end

  def is_appointment?
    type == 'IPM.Appointment'
  end

  def is_contact?
    type == 'IPM.Contact'
  end

  def is_meeting_request?
    type == 'IPM.Schedule.Meeting.Request'
  end

  def is_meeting_cancelled?
    type == 'IPM.Schedule.Meeting.Canceled'
  end

  def is_non_delivery_report?
    type == 'REPORT.IPM.Note.NDR'
  end

  def get_sender(email)
    return nil if not_email?
    if email.SenderEmailType == 'EX'
      email.sender.GetExchangeDistributionList&.PrimarySmtpAddress || email.sender.GetExchangeUser&.PrimarySmtpAddress
    else
      email.senderemailaddress
    end
  rescue
    nil # i.e cancelled calendar event AND "Undeliverable: bla bla" emails
  end

  def dump_attachments
    Dir.mkdir('attachments') if !File.exists?('attachments')
    raw.attachments.Count.downto(1) do |i|
      raw.attachments.Item(i).saveasfile("#{File.expand_path(File.dirname(__FILE__))}/attachments/#{raw.creationtime.strftime('%Y%m%d %H%M%s')}.jpeg")
    end
  end

  def subject
    @subject
  end

  def sender
    @sender
  end

  def body
    @body
  end

  def age_days
    @age_days
  end
  
  def raw
    @raw
  end

  def delete
    raw.delete
  end

  def move(folder)
    raw.move(folder)
  end

  def empty?
    @body.split("\r\n").empty?
  end

  def created_at
    @created_at
  end
end