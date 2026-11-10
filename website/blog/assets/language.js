document.querySelectorAll('.languages a').forEach(function(a){a.addEventListener('click',function(){a.href=a.getAttribute('href').split('#')[0]+location.hash;});});
