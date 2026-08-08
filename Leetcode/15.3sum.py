class Solution:
    def threeSum(self, nums: list[int]) -> list[list[int]]:
        n = len(nums)
        nums.sort()
        ans = []
        for i in range(0, n-2):
            if i != 0 and nums[i] == nums[i-1]:
                continue
            l = i+1
            r = n-1
            while l < r:
                sum_ = nums[i] + nums[l] + nums[r]
                if sum_ == 0:
                    ans.append([nums[i], nums[l], nums[r]])
                    l += 1
                    while l < r and nums[l] == nums[l-1]:
                        l += 1
                elif sum_ < 0:
                    l += 1
                else:
                    r -= 1

        return ans  # outside the for loop
