# Plugin's routes
# See: http://guides.rubyonrails.org/routing.html
get 'import_issue', :to => 'excel_sheet#index'
post 'upload_sheet', :to => 'excel_sheet#upload_sheet'
