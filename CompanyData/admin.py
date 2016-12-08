from django.contrib import admin
from CompanyData.models import *

class CompanyAdmin(admin.ModelAdmin):	
	search_fields = ['companyid']

admin.site.register(CompanyData, CompanyAdmin)

