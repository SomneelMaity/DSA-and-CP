/*
 * @lc app=leetcode id=20 lang=cpp
 *
 * [20] Valid Parentheses
 */

// @lc code=start
class Solution {
public:
    bool isValid(string s) {
        int n = s.size();
        if(n == 1) return false;
        stack<char> st;

        for(int i = 0; i < n; i++){
            if(s[i] == '(' || s[i] == '{' || s[i] == '['){
                st.push(s[i]);
            }else{
                if(st.empty()) return false;
                else if(s[i] == ')' and st.top() != '(') return false;
                else if(s[i] == '}' and st.top() != '{') return false;
                else if(s[i] == ']' and st.top() != '[') return false;
    }
    return true;
}
}
};
// @lc code=end

