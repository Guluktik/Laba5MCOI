using OfficeOpenXml;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.Statistics;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class HierarchicalClustering
{
    private Matrix<double> _data; // Исходные данные
    private Matrix<double> _distanceMatrix; // Матрица расстояний
    private List<List<int>> _clusters; // Список кластеров (список списков индексов)

    public HierarchicalClustering(Matrix<double> data)
    {
        _data = data;
        _distanceMatrix = CalculateDistanceMatrix(data); // Инициализация матрицы расстояний
        _clusters = InitializeClusters(data.RowCount); // Инициализация кластеров
    }

    // Инициализация: каждый объект — свой кластер
    private List<List<int>> InitializeClusters(int numObjects)
    {
        var clusters = new List<List<int>>();
        for (int i = 0; i < numObjects; i++)
        {
            clusters.Add(new List<int> { i });
        }
        return clusters;
    }

    // Расчет матрицы расстояний (евклидово расстояние)
    public Matrix<double> CalculateDistanceMatrix(Matrix<double> matrix)
    {
        int n = matrix.RowCount;
        var distanceMatrix = Matrix<double>.Build.Dense(n, n);

        for (int i = 0; i < n; i++)
        {
            for (int j = i + 1; j < n; j++)
            {
                double distance = (matrix.Row(i) - matrix.Row(j)).L2Norm();
                distanceMatrix[i, j] = distance;
                distanceMatrix[j, i] = distance;
            }
        }
        return distanceMatrix;
    }

    // Расчет центра тяжести для заданного кластера
    private Vector<double> CalculateCentroid(List<int> cluster)
    {
        var centroid = Vector<double>.Build.Dense(_data.ColumnCount);
        foreach (int index in cluster)
        {
            centroid += _data.Row(index);
        }
        return centroid / cluster.Count;
    }

    // Поиск двух ближайших кластеров
    private (int, int) FindClosestClusters()
    {
        double minDistance = double.MaxValue;
        int cluster1 = -1;
        int cluster2 = -1;

        for (int i = 0; i < _clusters.Count; i++)
        {
            for (int j = i + 1; j < _clusters.Count; j++)
            {
                double distance = CalculateClusterDistance(_clusters[i], _clusters[j]);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    cluster1 = i;
                    cluster2 = j;
                }
            }
        }
        return (cluster1, cluster2);
    }

    // Расчет расстояния между кластерами по их центрам тяжести
    private double CalculateClusterDistance(List<int> cluster1, List<int> cluster2)
    {
        var centroid1 = CalculateCentroid(cluster1);
        var centroid2 = CalculateCentroid(cluster2);
        return (centroid1 - centroid2).L2Norm();
    }

    // Выполнение иерархической кластеризации
    public void PerformClustering()
    {
        int step = 1; // Для подсчета шагов объединения
        while (_clusters.Count > 1)
        {
            var (cluster1, cluster2) = FindClosestClusters();

            // Состав объединяемых кластеров для вывода
            var cluster1Elements = string.Join(", ", _clusters[cluster1]);
            var cluster2Elements = string.Join(", ", _clusters[cluster2]);

            // Расстояние между кластерами
            double distance = CalculateClusterDistance(_clusters[cluster1], _clusters[cluster2]);

            // Объединяем кластеры
            _clusters[cluster1].AddRange(_clusters[cluster2]);
            _clusters.RemoveAt(cluster2);

            // Сортировка элементов в объединенном кластере
            _clusters[cluster1] = _clusters[cluster1].OrderBy(x => x).ToList();

            // Новый состав объединенного кластера для вывода
            var mergedClusterElements = string.Join(", ", _clusters[cluster1]);

            // Вывод информации об объединении кластеров в нужном формате
            Console.WriteLine($"Кластер Cluster {{ {cluster1Elements} }} и кластер Cluster {{ {cluster2Elements} }} с дистанцией между ними {distance:F4} объединены в кластер Cluster {{ {mergedClusterElements} }}");

            step++;
        }
    }

    // Загрузка данных из Excel
    public static Matrix<double> LoadDataFromExcel(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            var matrix = Matrix<double>.Build.Dense(rowCount - 1, colCount);

            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    matrix[i - 2, j - 1] = Convert.ToDouble(worksheet.Cells[i, j].Value);
                }
            }
            return matrix;
        }
    }

    // Вычисление среднего и стандартного отклонения
    public static (Vector<double> meanVector, Vector<double> stdDevVector) CalculateStatistics(Matrix<double> matrix)
    {
        var meanVector = Vector<double>.Build.Dense(matrix.ColumnCount, j => matrix.Column(j).Mean());
        var stdDevVector = Vector<double>.Build.Dense(matrix.ColumnCount, j => matrix.Column(j).StandardDeviation());
        return (meanVector, stdDevVector);
    }

    // Стандартизация данных
    public static Matrix<double> StandardizeMatrix(Matrix<double> matrix, Vector<double> meanVector, Vector<double> stdDevVector)
    {
        var standardizedMatrix = matrix.Clone();
        for (int i = 0; i < standardizedMatrix.RowCount; i++)
        {
            for (int j = 0; j < standardizedMatrix.ColumnCount; j++)
            {
                standardizedMatrix[i, j] = (standardizedMatrix[i, j] - meanVector[j]) / stdDevVector[j];
            }
        }
        return standardizedMatrix;
    }

    // Вычисление среднего квадратического отклонения
    public static Vector<double> CalculateMeanSquareDeviation(Matrix<double> matrix, Vector<double> meanVector)
    {
        int columns = matrix.ColumnCount;
        var meanSquareDeviation = Vector<double>.Build.Dense(columns);

        for (int j = 0; j < columns; j++)
        {
            double mean = meanVector[j]; // Используем ранее рассчитанное среднее значение
            double sumOfSquaredDeviations = matrix.Column(j).Select(x => Math.Pow(x - mean, 2)).Sum();
            meanSquareDeviation[j] = Math.Sqrt(sumOfSquaredDeviations / matrix.RowCount);
        }

        return meanSquareDeviation;
    }

    // Вывод вектора средних квадратических отклонений
    public static void PrintMeanSquareDeviation(Vector<double> meanSquareDeviation)
    {
        Console.WriteLine("Вектор средних квадратических отклонений:");
        foreach (var value in meanSquareDeviation)
        {
            Console.Write($"{value:F4}\t");
        }
        Console.WriteLine();
    }
    
}

class Program
{
    // Вывод матрицы расстояний
    public static void PrintDistanceMatrix(Matrix<double> distanceMatrix)
    {
        Console.WriteLine("Матрица расстояний:");
        for (int i = 0; i < distanceMatrix.RowCount; i++)
        {
            for (int j = 0; j < distanceMatrix.ColumnCount; j++)
            {
                Console.Write($"{distanceMatrix[i, j]:F4}\t");
            }
            Console.WriteLine();
        }
    }

    static void Main()
    {
        // Загрузка данных из Excel
        string filePath = @"D:\Тест.xlsx";
        var data = HierarchicalClustering.LoadDataFromExcel(filePath);

        // Вывод исходной матрицы
        Console.WriteLine("Исходная матрица:");
        for (int i = 0; i < data.RowCount; i++)
        {
            for (int j = 0; j < data.ColumnCount; j++)
            {
                Console.Write($"{data[i, j]:F2} ");
            }
            Console.WriteLine();
        }

        // Расчет среднего и стандартного отклонения
        var (meanVector, stdDevVector) = HierarchicalClustering.CalculateStatistics(data);
        Console.WriteLine("Вектор средних:");
        Console.WriteLine(meanVector);

        // Расчет среднего квадратического отклонения с использованием вектора средних
        var meanSquareDeviation = HierarchicalClustering.CalculateMeanSquareDeviation(data, meanVector);
        HierarchicalClustering.PrintMeanSquareDeviation(meanSquareDeviation);

        // Стандартизация матрицы
        var standardizedData = HierarchicalClustering.StandardizeMatrix(data, meanVector, stdDevVector);
        Console.WriteLine("Стандартизованная матрица:");
        for (int i = 0; i < standardizedData.RowCount; i++)
        {
            for (int j = 0; j < standardizedData.ColumnCount; j++)
            {
                Console.Write($"{standardizedData[i, j]:F2} ");
            }
            Console.WriteLine();
        }
        // Расчет и вывод матрицы расстояний
        var distanceMatrix = new HierarchicalClustering(data).CalculateDistanceMatrix(data);
        PrintDistanceMatrix(distanceMatrix);

        // Иерархическая кластеризация
        var clustering = new HierarchicalClustering(data);
        clustering.PerformClustering();
    }
}
