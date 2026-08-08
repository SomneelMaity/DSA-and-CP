/*
 * @lc app=leetcode id=64 lang=cpp
 *
 * [64] Minimum Path Sum
 *
 * https://leetcode.com/problems/minimum-path-sum/description/
 *
 * algorithms
 * Medium (68.65%)
 * Likes:    13836
 * Dislikes: 204
 * Total Accepted:    2M
 * Total Submissions: 2.8M
 * Testcase Example:  '[[1,3,1],[1,5,1],[4,2,1]]'
 *
 * Given a m x n grid filled with non-negative numbers, find a path from top
 * left to bottom right, which minimizes the sum of all numbers along its
 * path.
 * 
 * Note: You can only move either down or right at any point in time.
 * 
 * 
 * Example 1:
 * 
 * 
 * Input: grid = [[1,3,1],[1,5,1],[4,2,1]]
 * Output: 7
 * Explanation: Because the path 1 → 3 → 1 → 1 → 1 minimizes the sum.
 * 
 * 
 * Example 2:
 * 
 * 
 * Input: grid = [[1,2,3],[4,5,6]]
 * Output: 12
 * 
 * 
 * 
 * Constraints:
 * 
 * 
 * m == grid.length
 * n == grid[i].length
 * 1 <= m, n <= 200
 * 0 <= grid[i][j] <= 200
 * 
 * 
 */

// @lc code=start
class Solution {
    vector<vector<int>> memo;

    int solve(vector<vector<int>>& grid, int i, int j) {
        if (i == 0 && j == 0) return grid[0][0];
        if (memo[i][j] != -1) return memo[i][j];

        if (i == 0) return memo[i][j] = solve(grid, i, j-1) + grid[i][j];
        if (j == 0) return memo[i][j] = solve(grid, i-1, j) + grid[i][j];

        return memo[i][j] = min(solve(grid, i-1, j), solve(grid, i, j-1)) + grid[i][j];
    }

public:
    int minPathSum(vector<vector<int>>& grid) {
        int m = grid.size(), n = grid[0].size();
        memo.assign(m, vector<int>(n, -1));
        return solve(grid, m-1, n-1);
    }
};
// @lc code=end

