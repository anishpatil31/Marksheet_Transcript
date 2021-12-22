from django import forms
from django.db.models import fields
from .models import *


class TranscriptForm(forms.ModelForm):
    class Meta:
        model = RangeInput
        fields = ['left', 'right', 'stamp', 'sign']
