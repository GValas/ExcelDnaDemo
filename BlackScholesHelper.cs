using System;
using System.Linq;
using MathNet.Numerics.Distributions;

namespace ExcelDnaDemo
{
    public class BlackScholesHelper
    {

        private readonly Normal _rand = new Normal();
        private readonly int MclPaths = 10000000;

        public double CallPrice(
            double spot,
            double vol,
            double rate,
            double timeToMaturity,
            double strike)
        {

            // aliases
            var s0 = spot;
            var k = strike;
            var dt = timeToMaturity;
            var df = Math.Exp(-rate * dt);
            var vdt = vol * Math.Sqrt(dt);

            // montecarlo simulation price
            return Enumerable.Range(0, MclPaths)
                .Select(i => s0 * Math.Exp((rate * dt - vdt * vdt / 2 + _rand.Sample() * vdt)))
                .Select(s => Math.Max(s - k, 0))
                .Average() * df;
        }


    }
}
