Redmine::Plugin.register :issue_importer_xls do
  name 'Issue Importer Xls plugin'
  author 'Guruprasad Zapate'
  description 'Import Excel Sheet to create Redmine issues'
  version '0.0.1'


  permission :excel_sheet, { :excel_sheet => [:index, :upload_sheet] }, :public => true
  # menu :project_menu, :polls, { :controller => 'polls', :action => 'index' }, :caption => 'Polls', :after => :activity, :param => :project_id

  # menu :application_menu, :issue_importer_xls, { :controller => 'excel_sheet', :action => 'index' }, 
  # 			:caption => 'Import Issues' ,:last => true
  menu :project_menu, :excel_sheet, { :controller => 'excel_sheet', :action => 'index' }, 
  			:caption => 'Import Issues' ,:last => true

  settings :default => {'empty' => true}, :partial => 'settings/issue_importer_setting'
        
end
