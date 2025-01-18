// Импорт необходимых классов для работы с вводом/выводом, коллекциями и Apache POI для работы с Excel
import java.io.*;
import java.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xddf.usermodel.chart.*;

// Класс для хранения состояния игры, реализует интерфейс Serializable для сохранения в файл
class GameState implements Serializable {
    int[][] gameBoard; // Игровое поле
    int playerRice = 10; // Ресурсы игрока
    int playerWater = 10;
    int playerFarmers = 3;
    int aiRice = 10; // Ресурсы ИИ
    int aiWater = 10;
    int aiFarmers = 3;
    int playerTerritories = 1; // Территории игрока и ИИ
    int aiTerritories = 1;

    // История изменений ресурсов для отслеживания динамики
    List<Integer> playerRiceHistory = new ArrayList<>();
    List<Integer> playerWaterHistory = new ArrayList<>();
    List<Integer> playerFarmersHistory = new ArrayList<>();
    List<Integer> aiRiceHistory = new ArrayList<>();
    List<Integer> aiWaterHistory = new ArrayList<>();
    List<Integer> aiFarmersHistory = new ArrayList<>();

    // Конструктор состояния игры, инициализирует игровое поле
    GameState(int size) {
        gameBoard = new int[size][size];
        for (int[] row : gameBoard) {
            Arrays.fill(row, 2); // Заполнение начальными значениями
        }
        // Установка начальных позиций игрока и ИИ
        gameBoard[size - 1][size - 1] = -1;
        gameBoard[0][0] = -2;
    }

    // Метод для сохранения истории ресурсов
    void saveHistory() {
        playerRiceHistory.add(playerRice);
        playerWaterHistory.add(playerWater);
        playerFarmersHistory.add(playerFarmers);
        aiRiceHistory.add(aiRice);
        aiWaterHistory.add(aiWater);
        aiFarmersHistory.add(aiFarmers);
    }
}

// Класс для управления игрой
class Game {
    private GameState gameState;
    private final Scanner scanner = new Scanner(System.in); // Сканер для ввода данных пользователем

    // Метод для загрузки или начала новой игры
    void loadOrNewGame() {
        try {
            System.out.println("1. Новая игра\n2. Загрузить игру");
            int choice = scanner.nextInt();
            scanner.nextLine();

            if (choice == 2) {
                loadGame(); // Загрузка сохраненной игры
            } else {
                System.out.println("Введите размер игрового поля (например, 5 для 5x5): ");
                int size = scanner.nextInt();
                if (size <= 0) throw new IllegalArgumentException("Размер поля должен быть больше нуля.");
                scanner.nextLine();
                gameState = new GameState(size); // Создание нового состояния игры
            }
        } catch (InputMismatchException e) {
            System.out.println("Ошибка: Введите числовое значение.");
            scanner.nextLine();
        } catch (Exception e) {
            System.out.println("Ошибка: " + e.getMessage());
        }
    }

    // Метод для загрузки игры из файла
    void loadGame() {
        try (ObjectInputStream ois = new ObjectInputStream(new FileInputStream("game.save"))) {
            gameState = (GameState) ois.readObject(); // Десериализация состояния игры
            System.out.println("Игра загружена!");
        } catch (IOException | ClassNotFoundException e) {
            System.out.println("Ошибка загрузки игры: " + e.getMessage());
            System.exit(1); // Выход из программы при ошибке загрузки
        }
    }

    // Метод для сохранения игры в файл
    void saveGame() {
        try (ObjectOutputStream oos = new ObjectOutputStream(new FileOutputStream("game.save"))) {
            oos.writeObject(gameState); // Сериализация состояния игры
            System.out.println("Игра сохранена!");
        } catch (IOException e) {
            System.out.println("Ошибка сохранения игры: " + e.getMessage());
        }
    }

    // Метод для отображения игрового поля
    void displayGameBoard() {
        try {
            System.out.println("Текущее состояние карты:");
            System.out.print("   ");
            for (int i = 0; i < gameState.gameBoard.length; i++) {
                System.out.print(i + " ");
            }
            System.out.println();
            for (int i = 0; i < gameState.gameBoard.length; i++) {
                System.out.print(i + " |");
                for (int j = 0; j < gameState.gameBoard[i].length; j++) {
                    if (gameState.gameBoard[i][j] == -1) {
                        System.out.print("P "); // Позиция игрока
                    } else if (gameState.gameBoard[i][j] == -2) {
                        System.out.print("A "); // Позиция ИИ
                    } else if (gameState.gameBoard[i][j] == 0) {
                        System.out.print("X "); // Свободное поле
                    } else {
                        System.out.print(gameState.gameBoard[i][j] + " "); // Значение на поле
                    }
                }
                System.out.println();
            }
            System.out.println();
        } catch (Exception e) {
            System.out.println("Ошибка отображения карты: " + e.getMessage());
        }
    }

    // Метод для хода игрока
    void playerTurn() {
        try {
            System.out.println("Ваши ресурсы: Рис=" + gameState.playerRice + ", Вода=" + gameState.playerWater + ", Крестьяне=" + gameState.playerFarmers);
            System.out.println("Выберите действие:\n1. Набрать воды\n2. Полить рис\n3. Освоить территорию\n4. Построить дом крестьянина\n5. Сохранить игру\n6. Сохранить показатели в таблицу");
            int choice = scanner.nextInt();
            scanner.nextLine();

            switch (choice) {
                case 1 -> gameState.playerWater += 5; // Добавление воды
                case 2 -> {
                    if (gameState.playerWater > 0) {
                        gameState.playerWater--; // Уменьшение воды
                        gameState.playerRice += 3; // Увеличение запасов риса
                    } else {
                        System.out.println("Недостаточно воды!");
                    }
                }
                case 3 -> conquerTerritory(true); // Освоение территории игроком
                case 4 -> {
                    if (gameState.playerRice >= 5 && gameState.playerWater >= 3 && gameState.playerFarmers > 0) {
                        gameState.playerRice -= 5; // Расход риса
                        gameState.playerWater -= 3; // Расход воды
                        gameState.playerFarmers++; // Увеличение количества крестьян
                    } else {
                        System.out.println("Недостаточно ресурсов или крестьян!");
                    }
                }
                case 5 -> saveGame(); // Сохранение игры
                case 6 -> saveHistoryToExcel(); // Сохранение истории в Excel
                default -> System.out.println("Неверный выбор. Попробуйте снова.");
            }
        } catch (InputMismatchException e) {
            System.out.println("Ошибка ввода! Ожидается число.");
            scanner.nextLine(); // Очистка буфера ввода
        }
    }

    // Метод для хода ИИ
    void aiTurn() {
        try {
            Random random = new Random();
            int choice = random.nextInt(4) + 1; // Случайный выбор действия ИИ

            switch (choice) {
                case 1 -> gameState.aiWater += 5; // Добавление воды ИИ
                case 2 -> {
                    if (gameState.aiWater > 0) {
                        gameState.aiWater--; // Уменьшение воды у ИИ
                        gameState.aiRice += 3; // Увеличение запасов риса у ИИ
                    }
                }
                case 3 -> conquerTerritory(false); // Освоение территории ИИ
                case 4 -> {
                    if (gameState.aiRice >= 5 && gameState.aiWater >= 3 && gameState.aiFarmers > 0) {
                        gameState.aiRice -= 5; // Расход риса ИИ
                        gameState.aiWater -= 3; // Расход воды ИИ
                        gameState.aiFarmers++; // Увеличение количества крестьян у ИИ
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("Ошибка в ходе ИИ: " + e.getMessage());
        }
    }

    // Метод для завоевания территории
    void conquerTerritory(boolean isPlayer) {
        try {
            int x, y; // Координаты для освоения
            if (isPlayer) {
                System.out.println("Введите координаты для освоения (x y):");
                x = scanner.nextInt();
                y = scanner.nextInt();
            } else {
                Random random = new Random();
                x = random.nextInt(gameState.gameBoard.length); // Случайная координата x для ИИ
                y = random.nextInt(gameState.gameBoard.length); // Случайная координата y для ИИ
            }

            if (x < 0 || y < 0 || x >= gameState.gameBoard.length || y >= gameState.gameBoard.length) {
                throw new IndexOutOfBoundsException("Координаты вне диапазона!"); // Проверка корректности координат
            }

            if (gameState.gameBoard[x][y] > 0 && (isPlayer ? gameState.playerFarmers : gameState.aiFarmers) >= gameState.gameBoard[x][y]) {
                if (isPlayer) {
                    gameState.playerFarmers -= gameState.gameBoard[x][y]; // Уменьшение количества крестьян у игрока
                    gameState.playerTerritories++; // Увеличение территорий игрока
                    gameState.gameBoard[x][y] = -1; // Отметка территории как завоеванной игроком
                } else {
                    gameState.aiFarmers -= gameState.gameBoard[x][y]; // Уменьшение количества крестьян у ИИ
                    gameState.aiTerritories++; // Увеличение территорий ИИ
                    gameState.gameBoard[x][y] = -2; // Отметка территории как завоеванной ИИ
                }
            } else if (isPlayer) {
                System.out.println("Недостаточно крестьян или некорректные координаты!");
            }
        } catch (InputMismatchException e) {
            System.out.println("Ошибка ввода! Попробуйте снова.");
            scanner.nextLine(); // Очистка буфера ввода
        } catch (IndexOutOfBoundsException e) {
            System.out.println("Ошибка: " + e.getMessage());
        }
    }

    // Метод для сохранения истории игры в файл Excel
    void saveHistoryToExcel() {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("История игры"); // Создание листа в книге Excel

            Row header = sheet.createRow(0);
            String[] columns = {"День", "Рис игрока", "Вода игрока", "Крестьяне игрока", "Рис ИИ", "Вода ИИ", "Крестьяне ИИ"};
            for (int i = 0; i < columns.length; i++) {
                header.createCell(i).setCellValue(columns[i]); // Заполнение заголовков столбцов
            }

            int days = gameState.playerRiceHistory.size();
            for (int i = 0; i < days; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(i + 1); // День игры
                row.createCell(1).setCellValue(gameState.playerRiceHistory.get(i)); // История риса игрока
                row.createCell(2).setCellValue(gameState.playerWaterHistory.get(i)); // История воды игрока
                row.createCell(3).setCellValue(gameState.playerFarmersHistory.get(i)); // История крестьян игрока
                row.createCell(4).setCellValue(gameState.aiRiceHistory.get(i)); // История риса ИИ
                row.createCell(5).setCellValue(gameState.aiWaterHistory.get(i)); // История воды ИИ
                row.createCell(6).setCellValue(gameState.aiFarmersHistory.get(i)); // История крестьян ИИ
            }

            for (int i = 0; i < columns.length; i++) {
                sheet.autoSizeColumn(i); // Автоматическая настройка ширины столбцов
            }

            try (FileOutputStream fileOut = new FileOutputStream("история_игры.xlsx")) {
                workbook.write(fileOut); // Сохранение файла Excel
            }
            System.out.println("История игры сохранена в файл история_игры.xlsx");
        } catch (IOException e) {
            System.out.println("Ошибка сохранения в Excel: " + e.getMessage());
        }
    }

    // Основной метод программы
    void startGame() {
        try {
            System.out.println("Добро пожаловать в игру!");
            loadOrNewGame(); // Загрузка или создание новой игры
            while (true) {
                try {
                    displayGameBoard(); // Отображение игрового поля
                    gameState.saveHistory(); // Сохранение истории
                    playerTurn(); // Ход игрока
                    aiTurn(); // Ход ИИ
                    endOfDay(); // Обработка событий конца дня
                    checkVictory(); // Проверка условий победы
                } catch (InputMismatchException e) {
                    System.out.println("Ошибка ввода! Попробуйте снова.");
                    scanner.nextLine(); // Очистка буфера ввода
                }
            }
        } catch (Exception e) {
            System.out.println("Критическая ошибка: " + e.getMessage());
            e.printStackTrace();
        }
    }

    // Метод для обработки событий конца дня (пустой метод, можно добавить логику)
    void endOfDay() {
    }

    // Метод для проверки условий победы
    void checkVictory() {
        try {
            int totalTerritories = gameState.gameBoard.length * gameState.gameBoard.length; // Общее количество территорий
            if (gameState.playerTerritories >= totalTerritories / 2) {
                System.out.println("Вы победили!"); // Победа игрока
                saveHistoryToExcel(); // Сохранение истории в Excel
                System.exit(0); // Завершение программы
            } else if (gameState.aiTerritories >= totalTerritories / 2) {
                System.out.println("ИИ победил!"); // Победа ИИ
                saveHistoryToExcel(); // Сохранение истории в Excel
                System.exit(0); // Завершение программы
            }
        } catch (Exception e) {
            System.out.println("Ошибка проверки победы: " + e.getMessage());
        }
    }
}

public class MVPGame {
    public static void main(String[] args) {
        Game game = new Game();
        game.startGame();
    }
}