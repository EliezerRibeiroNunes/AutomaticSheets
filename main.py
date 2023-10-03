from controllers.main_controller import MainController
from views.main_view import MainView

if __name__ == "__main__":
    controller = MainController()
    view = MainView(controller)

