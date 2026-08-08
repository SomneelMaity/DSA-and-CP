/*
 * @lc app=leetcode id=110 lang=cpp
 *
 * [110] Balanced Binary Tree
 *
 * https://leetcode.com/problems/balanced-binary-tree/description/
 *
 * algorithms
 * Easy (58.88%)
 * Likes:    12353
 * Dislikes: 848
 * Total Accepted:    2.6M
 * Total Submissions: 4.4M
 * Testcase Example:  '[3,9,20,null,null,15,7]'
 *
 * Given a binary tree, determine if it is height-balanced.
 * 
 * 
 * Example 1:
 * 
 * 
 * Input: root = [3,9,20,null,null,15,7]
 * Output: true
 * 
 * 
 * Example 2:
 * 
 * 
 * Input: root = [1,2,2,3,3,null,null,4,4]
 * Output: false
 * 
 * 
 * Example 3:
 * 
 * 
 * Input: root = []
 * Output: true
 * 
 * 
 * 
 * Constraints:
 * 
 * 
 * The number of nodes in the tree is in the range [0, 5000].
 * -10^4 <= Node.val <= 10^4
 * 
 * 
 */

// @lc code=start
/**
 * Definition for a binary tree node.
 * struct TreeNode {
 *     int val;
 *     TreeNode *left;
 *     TreeNode *right;
 *     TreeNode() : val(0), left(nullptr), right(nullptr) {}
 *     TreeNode(int x) : val(x), left(nullptr), right(nullptr) {}
 *     TreeNode(int x, TreeNode *left, TreeNode *right) : val(x), left(left), right(right) {}
 * };
 */
class Solution {
    // returns height, or -1 if subtree is unbalanced
    int check(TreeNode* root) {
        if (!root) return 0;
        int l = check(root->left);
        if (l == -1) return -1;
        int r = check(root->right);
        if (r == -1) return -1;
        if (abs(l - r) > 1) return -1;
        return 1 + max(l, r);
    }

public:
    bool isBalanced(TreeNode* root) {
        return check(root) != -1;
    }
};
// @lc code=end

