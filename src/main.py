import Model
import View
import Controller


def main():
    model = Model.Model()
    view = View.View()
    Controller.Controller(model, view)






if __name__ == "__main__":
    main()