from django.shortcuts import render

# Create your views here.

def research_dashboard(request):
    return render(request, 'research/dashboard.html')