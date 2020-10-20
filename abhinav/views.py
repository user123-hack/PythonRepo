from django.shortcuts import render
#from django.http import HttpResponse
from .models import Post

'''posts = [
    {
        'author': 'coreyMs',
        'title' : 'Blog Post 1',
        'content' : 'first post content',
        'date_posted' : 'March 28,2020'
    },
    {
        'author': 'abhinav',
        'title' : 'Blog Post 2',
        'content' : 'first post content for blog',
        'date_posted' : 'March 28,2020'
    }
]Dont need now'''

def home(request):
    context = {
        'posts' : Post.objects.all()
    }
    return render(request, 'blog/home.html', context)

def about(request):
    return render(request, 'blog/about.html', {'title' : 'About'})

