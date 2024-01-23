# Function to print union of two unsorted arrays


def Union(a, b, n, m):
    result = [0 for _ in range(n + m)]

    index, left, right = 0, 0, 0
    while left < n and right < m:
        print(a[left], b[right])
        if a[left] < b[right]:
            if index != 0 and a[left] == result[index - 1]:
                left += 1
            else:
                result[index] = a[left]
                left += 1
                index += 1

        else:
            if index != 0 and b[right] == result[index - 1]:
                right += 1
            else:
                result[index] = b[right]
                right += 1
                index += 1

    while left < n:
        if index != 0 and a[left] == result[index - 1]:
            left += 1
        else:
            result[index] = a[left]
            left += 1
            index += 1

    while right < m:
        if index != 0 and b[right] == result[index - 1]:
            right += 1
        else:
            result[index] = b[right]
            right += 1
            index += 1

    print("Union:", *result[:index])


Union([1, 2, 3, 4], [1, 1, 2, 3, 4], 4, 5)
