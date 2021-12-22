from django.shortcuts import redirect, render
from .forms import *
from .models import *
# Create your views here.
def HomePage(request) :
    if request.method == 'POST' :
        form = TranscriptForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            if request.POST.get('all'):
                return redirect('Proj2')
            if request.POST.get('range'):
                return redirect('range')
    else :
        form = TranscriptForm()
    return render(request, 'polls/HomePage.html', {'form': form, 'pop_list':[], 'f':0})