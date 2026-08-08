/*
 * @lc app=leetcode id=438 lang=cpp
 *
 * [438] Find All Anagrams in a String
 */

// @lc code=start
class Solution {
public:
    vector<int> findAnagrams(string s, string p) {
        int n = s.size(), m = p.size();
        if(m>n) return {};
        unordered_map<char, int> mapp1, mapp2;

        for(char c: p)
            mapp1[c]++;
        
        for(int i=0; i<m; i++)
        {
            mapp2[s[i]]++;
        }
        vector<int> ans;
        if(mapp1 == mapp2) ans.push_back(0);

        int l = 0;
        for(int r=m; r<n; r++)
        {
            mapp2[s[r]]++;
            mapp2[s[l]]--;
            if(mapp2[s[l]] == 0) mapp2.erase(s[l]);
            l++;
            if(mapp1 == mapp2) ans.push_back(l);
        }
        return ans;

    }
};
// @lc code=end

