def get_matrix(ws):
    matrix = []
    for i, row in enumerate(ws.rows):
        if i == 0:
            continue
        matrix.append([])
        for j in range(len(row)):
            if j == 0:
                continue
            matrix[i - 1].append(row[j].value)

    return matrix


def get_head(ws):
    head = [cell for cell in next(ws.iter_rows(min_row=1,
                                               min_col=2,
                                               max_row=1,
                                               max_col=ws.max_column,
                                               values_only=True))]
    return head


def get_wall(ws):
    wall = []
    for row in ws.rows:
        if row[0].value is not None:
            wall.append(row[0].value)
    return wall