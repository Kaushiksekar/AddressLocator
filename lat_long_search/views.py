from django.http import HttpResponse
from django.template import loader
from .forms import UploadFileForm
import xlsxwriter
import xlrd
import geocoder

# Create your views here.
def index(request):
	template = loader.get_template("lat_long_search/index.html")
	return HttpResponse(template.render({}, request))

def get_excel(request):
	if request.method == "POST":
		form = UploadFileForm(request.POST, request.FILES)
		if form.is_valid():
			file_name = request.FILES["excel_file"].name
			if not is_accepted_extension(file_name):
				output1 = "Incorrect File extension"
				return HttpResponse(output1)
			else:
				output1 = read_excel_file(request.FILES['excel_file'])
			return output1         
		else:
			output1 = "Form invalid"
	return HttpResponse(output1)

def is_accepted_extension(file_name):
	extension = file_name.split(".")[-1]
	if extension == "xlsx" or extension == "xls":
		return True
	return False

def read_excel_file(excel_file):
	with open("lat_long_search/media/temp.xlsx", "wb+") as dest:
		for chunk in excel_file.chunks():
			dest.write(chunk) 
	book = xlrd.open_workbook("lat_long_search/media/temp.xlsx")
	ws = book.sheet_by_index(0)
	list1 = []
	for i in range(ws.nrows):
		list1.append(ws.cell(i, 0).value)
	list2, list3 = get_output_list(list1)
	excel_file = write_excel_file_xlsx(list1, list2, list3, excel_file)
	response = download_excel(excel_file)
	return response

def get_output_list(list1):
	list2 = []
	list3 = []
	for i, item in enumerate(list1):
		location = geocoder.locationiq(item, key="9cc01a46febc8d")
		print("Item",str(i+1))
		lat_long_tuple = get_lat_long_str_from_json(location.json)
		list2.append(lat_long_tuple[0])
		list3.append(lat_long_tuple[1])
	return list2, list3

def write_excel_file_xlsx(list1, list2, list3, excel_file):
	wb = xlsxwriter.Workbook("lat_long_search/media/temp1.xlsx")
	text_format = wb.add_format({'text_wrap': True})
	ws = wb.add_worksheet()
	ws.write("A1", "Address", text_format)
	ws.write("B1", "Latitude", text_format)
	ws.write("C1", "Longitude", text_format)
	for i in range(len(list1)):
		ws.write("A"+str(i+2), list1[i], text_format)
		ws.write("B"+str(i+2), list2[i], text_format)
		ws.write("C"+str(i+2), list3[i], text_format)
	wb.close()
	return excel_file

def get_lat_long_str_from_json(json_dict):
	try:
		latitude = json_dict['lat']
		longitude = json_dict['lng']
	except Exception as e:
		latitude = str(e)
		longitude = str(e)
	return str(latitude), str(longitude)

def download_excel(excel_file):
	download_file = open("lat_long_search/media/temp1.xlsx", "rb")
	response = HttpResponse(content=download_file, content_type="application/vnd.ms-excel")
	response['Content-Disposition'] = 'inline; filename=output.xlsx'
	return response