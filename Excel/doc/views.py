from django.shortcuts import render

# Create your views here.

def doc(request):
    return render(template_name='doc.html', request=request)