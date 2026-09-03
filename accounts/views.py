from django.shortcuts import redirect, render
from django.http import HttpResponse
from django.contrib.auth.models import User
from django.contrib.auth import authenticate, login, logout


def signup_view(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        email = request.POST.get('email')
        password = request.POST.get('password')

        if User.objects.filter(username=username).exists():
            return HttpResponse("""
                <h2>Username already exists!</h2>
                <a href="/accounts/signup/">Try again</a>
            """)

        user = User.objects.create_user(username=username, email=email, password=password)
        login(request, user)
        return redirect('/papers/')

    return render(request, 'signup.html')


def login_view(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')

        user = authenticate(request, username=username, password=password)

        if user is not None:
            login(request, user)
            return redirect('/papers/')
        else:
            return HttpResponse("""
                <h2>Invalid username or password!</h2>
                <a href="/accounts/login/">Try again</a>
            """)

    return render(request, 'login.html')


def logout_view(request):
    logout(request)
    return redirect('/accounts/login/')