class ExcelSheetController < ApplicationController
  unloadable
  	before_filter :find_project, :authorize, :only => :index

	def index

		@project = Project.find(params[:id])
		logger.info "inside ExcelSheetController"
		logger.info "Project ID is #{@project.id}"
		session[:project_id]=params[:id]
		
	end

	def save_configuration

	end

	def upload_sheet

		   logger.info "Inside upload_sheet method"
		   uploaded_io = params[:file]

		   if uploaded_io.nil? || uploaded_io.tempfile.nil?
		   		logger.info " No file found"
		   		flash[:notice] = 'Please Submit Excel File'
  				redirect_to :action => 'index', :id => session[:project_id]
  				return
		   end
		  
		   unless Dir.exists?("#{Rails.root}/public/uploads") 
		   		Dir::mkdir("#{Rails.root}/public/uploads")
		   end
		 
		   logger.info "File size  is #{uploaded_io.tempfile.size}"
		   logger.info "File path is #{uploaded_io.tempfile.to_path.to_s}"

		  FileUtils.cp  "#{uploaded_io.tempfile.to_path.to_s}", "#{Rails.root}/public/uploads/#{uploaded_io.original_filename}"		 
		
		 extname=File.extname("#{Rails.root}/public/uploads/#{uploaded_io.original_filename}")
		 logger.info "File extension #{extname}"

		 case extname
		 #Microsoft Excel File
		 when ".xls"
		 	workbook = Roo::Excel.new  "#{Rails.root}/public/uploads/#{uploaded_io.original_filename}"
		 #Microsoft Excel Xml File
		 when ".xlsx"
		 	workbook =  Roo::Excelx.new  "#{Rails.root}/public/uploads/#{uploaded_io.original_filename}"
		 #ODF Spreadsheet /OpenOffice document
		 when ".ods"
		 	workbook =Roo::OpenOffice.new   "#{Rails.root}/public/uploads/#{uploaded_io.original_filename}"
		 else
		 	logger.info "Please insert correct file"
		 	flash[:notice] = 'Please Submit Excel File'
  			redirect_to :action => 'index', :id => session[:project_id]
  			return
		 end
 		
 		 workbook.default_sheet = workbook.sheets[0]

 		 # logger.info " #{workbook.info} "


 		 headers = Hash.new
			workbook.row(1).each_with_index {|header,i|
				headers[header] = i
		 }
		 logger.info " #{headers}"

		 logger.info "Label ---------------- #{l(:label_import_issue_upload_file_note)}"
 		 #iterate over the file
 		 #get project name from first cell of the sheet

 		 project_name=workbook.cell(1,1)
 		 redmine_project = Project.find(session[:project_id])
 		 if !redmine_project
        	redmine_project = @redmine_project
      	 end
 		 logger.info " Current user #{User.current}"
 		 excel_error_message="Excel File contains following error.<br>"
 		 excel_having_errors=false

 		 ((workbook.first_row + 1)..workbook.last_row).each do |row|

 		 	row_content=Array.new(workbook.row(row))
 		 	##Row must contain task description to  create issue
 		 	if row_content[0].nil? 

 		 		excel_error_message.concat("Excel Row #{row} contains error.<br>")
 		 		excel_having_errors=true
 		 	end	

 		 end 



 		 unless  excel_having_errors

 		 logger.info "#{Setting.plugin_issue_importer_xls}"

 		 settings_conf=Setting.plugin_issue_importer_xls

 		 ((workbook.first_row + 1)..workbook.last_row).each do |row|

	 		 	#iterate through all rows
	 		 	row_content=Array.new(workbook.row(row))
	 		 	#Project Name/Task	Best Case	Worst Case	Average Case	Notes	Questions	Start Date	Due Date	Total(in weeks)	Asignee
	 		 	unless row_content[0]== l(:label_import_issue_task) || row_content[0] == l(:label_import_issue_design) || row_content[0] == l(:label_import_issue_development) || row_content[0] == l(:label_import_issue_documentation) || row_content[0] == l(:label_import_issue_testing) 

	 		 	    issue = Issue.new
				    issue.author_id = User.current.id
				 	issue.project_id = redmine_project.id
				 	issue.subject=row_content[settings_conf['task_column'].to_i]
				 	issue.tracker_id=2 #Feature
				 	issue.status_id=1 #New
				 	issue.description=row_content[4]
				 	issue.estimated_hours=row_content[settings_conf['average_hour_column'].to_i]
				 	issue.start_date=row_content[settings_conf['start_date_column'].to_i]
				 	issue.due_date=row_content[settings_conf['end_date_column'].to_i]

				 	User.all.each do |user|
				 		logger.info "----------------#{user.name} #{row_content[9]}------------------ "
				 		if user.name.eql? row_content[settings_conf['asignee_name_column'].to_i]
				 			issue.assigned_to_id=user.id
				 		end
				 	end
				 	#Save issue for project
			 		issue.save
	 			end  		

	 		 end 

	 	 else

	 	 	flash[:notice]=excel_error_message
	 	 	redirect_to :action => 'index', :id => session[:project_id]
	 	 	return
 		 	
 		 end

		
		flash[:notice] = 'Issues successfully created'
  		redirect_to :action => 'index', :id => session[:project_id]
	
	end

end
