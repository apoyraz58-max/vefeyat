class FramePolicyMiddleware:
    """Allow the Office task pane while protecting the editor and public site."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        if request.path != "/word/":
            response["X-Frame-Options"] = "DENY"
        return response
