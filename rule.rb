class Rule
  def initialize(params)
    @params = params
  end

  def applies?(email)
    #if email.subject.include?('Daily Status Email')
    #  binding.pry
    #end
    return false if (!@params[:type] && email.not_email?)

    (!@params[:body] || email.body =~ @params[:body]) &&
    (!@params[:subject] || email.subject =~ @params[:subject]) &&
    (!@params[:sender] || email.sender == @params[:sender]) &&
    (!@params[:sender_like] || email.sender =~ @params[:sender_like]) &&
    #(!@params[:to] || email.senderemailaddress == @params[:to]) &&
    (!@params[:older_than_days] || email.age_days > @params[:older_than_days]) &&
    (!@params[:type] || email.send(@params[:type]) == true)
  end

  def move_to
    @params[:move_to]
  end
end