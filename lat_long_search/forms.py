from django import forms

class UploadFileForm(forms.Form):
	excel_sfile = forms.FileField(required=False)