/*
 * @lc app=leetcode id=242 lang=cpp
 *
 * [242] Valid Anagram
 */

// @lc code=start
class Solution {
public:
    bool isAnagram(string s, string t) {
        if(s.size() != t.size()) return false;
        unordered_map<char, int> mapp1, mapp2;

        for(int i=0; i<s.size(); i++)
        {
            mapp1[s[i]]++;
            mapp2[t[i]]++;
        }

        return mapp1 == mapp2;
    }
};
// @lc code=end

