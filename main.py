"""
Arquivo principal do projeto Multiplica Evidências.

Este arquivo é responsável por preparar o ambiente da aplicação,
inicializar o banco de dados e abrir a interface gráfica.
"""


from config import ensure_project_dirs
from database import Database
from interface.app import MultiplicaApp


def main() -> None:
    ensure_project_dirs()
    database = Database()
    database.initialize()
    app = MultiplicaApp(database)
    app.mainloop()


if __name__ == "__main__":
    main()
