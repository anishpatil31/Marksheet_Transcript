from django import forms
from django.db.models import fields
from .models import *

class Markform(forms.ModelForm) :
    class Meta :
        model = Marktsheet_data
        fields = ['positive', 'negative', 'master_roll', 'response_csv']
        # fields = "__all__"



