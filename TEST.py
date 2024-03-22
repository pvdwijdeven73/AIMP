def FindMaxSum(arr, n):
    take = 0
    leave = 0
    for i in arr:
        prev_leave = leave
        leave = max(leave, take)
        take = prev_leave + i
    return max(leave, take)


arr = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
print(FindMaxSum(arr, len(arr)))
