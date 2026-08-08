/*
 * @lc app=leetcode id=75 lang=cpp
 *
 * [75] Sort Colors
 */

// @lc code=start
class Solution {
public:
    // Dutch National Flag: single pass, O(1) space
    void sortColors(vector<int>& nums) {
        int lo = 0, mid = 0, hi = nums.size() - 1;
        while (mid <= hi) {
            if (nums[mid] == 0)
                swap(nums[lo++], nums[mid++]);
            else if (nums[mid] == 1)
                mid++;
            else
                swap(nums[mid], nums[hi--]);
        }
    }
};
// @lc code=end

