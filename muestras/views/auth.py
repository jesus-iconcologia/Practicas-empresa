from django.contrib.auth.views import LoginView, LogoutView
from django.urls import reverse


class GesLabLoginView(LoginView):
    template_name = "registration/login.html"
    redirect_authenticated_user = True

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context["logged_out"] = self.request.GET.get("logged_out") == "1"
        return context


class GesLabLogoutView(LogoutView):
    http_method_names = ["post", "options"]

    def get_default_redirect_url(self):
        return f"{reverse('login')}?logged_out=1"
