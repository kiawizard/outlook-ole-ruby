require 'win32ole'

class OC; end

class Outlook
  def initialize
    ol = WIN32OLE.new('Outlook.Application')
    WIN32OLE.const_load(ol, OC)

    @mapi = ol.GetNameSpace("MAPI")
  end

  def raw
    @mapi
  end

  def inbox
    @inbox ||= @mapi.GetDefaultFolder(OC::OlFolderInbox)
  end

  #outlook.folder('Ivan.Kokorev@excelian.com/Inbox')
  def folder(path)
    current = raw
    path.split('/').each do |part|
      current = current.folders.item(part)
    end
    current
  end

  def autoclean_folder(path, delete_rules = [], move_rules = [])
    folder = folder(path)
    puts "#{folder.Items.Count} emails exist in #{path}"
    deleted = 0
    moved = 0
    folder.Items.Count.downto(1) do |i|
      puts "#{i} left to process, total: #{folder.Items.Count}" if i%100 == 0
      mail = Email.new(folder.Items(i))
      if delete_rules.any?{|dr| dr.applies?(mail)}
        mail.delete
        deleted += 1
      elsif (selected_rule = move_rules.detect{|dr| dr.applies?(mail)})
        mail.move(folder(selected_rule.move_to))
        moved += 1
      end
    end
    puts "Deleted #{deleted} emails"
    puts "Moved #{moved} emails"
  end

  def dump_folder_attachments_and_delete(path, params = {})
    folder = folder(path)
    folder.Items.Count.downto(1) do |i|
      puts "#{i} left to process, total: #{folder.Items.Count}" if i%100 == 0
      mail = Email.new(folder.Items(i))
      if !params[:empty_body_only] || mail.empty?
        mail.dump_attachments
        mail.delete
      end
    end
  end

  OlStoreType = {olStoreANSI: 3, olStoreDefault: 1, olStoreUnicode: 2}

  def open_or_create_datastore(name, file_path = nil, folder_name = 'Inbox')
    raw.folders.item(name)
  rescue
    raw.addstoreex(file_path, OlStoreType[:olStoreUnicode])
    store = raw.folders.item(raw.folders.Count)
    store.name = name
    store.folders.add(folder_name)
    store
  end

  def archive_mail(source, datastore_name_prefix, base_directory, params = {})
    params[:folder_name] ||= 'Inbox'
    params[:divide_by] ||= :year

    moved = 0
    folder = folder(source)
    folder.Items.Count.downto(1) do |i|
      puts "#{i} left to process, total: #{folder.Items.Count}" if i%100 == 0
      mail = Email.new(folder.Items(i))
      if !params[:rules] || params[:rules].any?{|rule| rule.applies?(mail)}
        store_name = "#{datastore_name_prefix} #{mail.created_at.year}#{" #{mail.created_at.month}" if params[:divide_by] == :month}"
        store_filename = "#{base_directory}/#{store_name}.pst"
        store = open_or_create_datastore(store_name, store_filename)
        target = store.folders.item('Inbox')
        mail.move(target)
        moved += 1
      end
    end

    puts "Moved #{moved} emails"
  end
end