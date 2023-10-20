# Best Metric Marker
This matlab code helps researchers to automatically mark the best metrics, and write the results in the Excel file. 

## Usage
Please read `best_metric_marker.m`, where there are clear descriptions about how to use this code. 

## Example 1
Suppose you have obtain the metrics of n remote sensing image
processing algorithms on different datasets, and the metric is
organized as follows:
|                 | dataset1     | dataset2    |
| :---:           | :---:        | :---:       |
| PSNR   of alg.1 |  p11         | p12         |
| SSIM   of alg.1 |  s11         | s12         |
| ERGAS  of alg.1 |  e11         | e12         |
| SAM    of alg.1 |  a11         | a12         |
| PSNR   of alg.2 |  p21         | p22         |
| SSIM   of alg.2 |  s21         | s22         |
| ERGAS  of alg.2 |  e21         | e22         |
| SAM    of alg.2 |  a21         | a22         |

where higher PSNR and SSIM lead to better results, and lower ERGAS
and SAM lead to better results. We want to round SSIM values as 3
digits, and round others as 2 digits. And mark the 1st and 2nd and
3rd best metrics as 'bold', 'color', 'italic'. The code should
be


```
data = [22.48	21.66	20.84	21.06	20.46	19.77 % PSNR  of alg.1
    0.524	0.456	0.386	0.449	0.403	0.358     % SSIM  of alg.1
    53.55	60.37	69.26	51.56	56.59	63.71     % ERGAS of alg.1
    13.96	14.68	15.48	13.33	14.06	14.99     % SAM  of alg.1
    22.68	21.98	21.26	22.32	21.71	21.05
    0.494	0.424	0.352	0.506	0.451	0.396
    52.11	56.95	61.79	43.4	46.93	51.35
    14.1	14.52	14.8	10.37	10.86	11.54
    32.57	30.86	28.38	30.92	29.01	26.32
    0.937	0.91	0.848	0.912	0.87	0.78
    15.86	19.28	25.65	15.67	19.47	26.43
    6.77	7.71	9.3	5.67	6.72	8.59
    33.58	32.02	30.11	31.44	29.83	27.87
    0.946	0.926	0.891	0.909	0.877	0.823
    14.64	17.23	21.13	15.38	18.13	22.22
    7.79	8.84	10.28	6.67	7.7	9.11
    ];

num_metrics = 4;
precision = [2,3,2,2];
optval = {'max', 'max', 'min', 'min'};
highlight.key = {'bold', 'color', 'italic'};
highlight.value = {true, [255,0,0], true};
border_mode = {'top', 'bottom', 'mid'};

best_metric_marker(data, ...
    num_metrics, ...
    'precision', precision, ...
    'optval', optval, ...
    'border_mode', border_mode, ...
    'highlight', highlight);
```

You will get the following table in the Excel file. 

 ![image](https://github.com/shuangxu96/best_metric_marker/blob/main/example1.jpg)

 
## Example 2

Suppose you have obtain AUC of n algorithms on different datasets,
and the metric is organized as follows:
|                 | dataset1     | dataset2    |
| :---:           | :---:        | :---:       |
| AUC   of alg.1 |  p11         | p12         |
| AUC   of alg.2 |  p21         | p22         |
| AUC   of alg.3 |  p31         | p32         |

where higher AUC leads to better results. We want to round AUC
values as 4 digits. And mark the 1st and 2nd best metrics as 'bold',
'underline'. The code should be


```
data = [0.9681	0.9698	0.9572	0.8447 % AUC  of alg.1
    0.9487	0.8786	0.9413	0.6845     % AUC  of alg.2
    0.9594	0.9179	0.9573	0.8401];   % AUC  of alg.3

num_metrics = 1;
precision = 4;
optval = {'max'};
highlight.key = {'bold', 'underline'};
highlight.value = {true,  true};
border_mode = {'top', 'bottom'};

best_metric_marker(data, ...
    num_metrics, ...
    'precision', precision, ...
    'optval', optval, ...
    'border_mode', border_mode, ...
    'highlight', highlight);
```

You will get the following table in the Excel file. 

 ![image](https://github.com/shuangxu96/best_metric_marker/blob/main/example2.png)
