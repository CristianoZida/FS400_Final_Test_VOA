
'''
Author(s):     Ruizhi Shi, Ran Ding, Yury Deshko, Michael Schmidt
imported from IMRA test code
'''
import numpy as np
from numpy import fft
import pylab as pl
import matplotlib
import matplotlib.pyplot as plt
import numpy, scipy.optimize
import pandas as pd

#----This is for plot and fit the signal to sin and calculate the phase----
def fit_sin(tt, yy):
    '''Fit sin to the input time sequence, and return fitting parameters "amp", "omega", "phase", "offset", "freq", "period" and "fitfunc"'''
    tt = numpy.array(tt)
    yy = numpy.array(yy)
    ff = numpy.fft.fftfreq(len(tt), (tt[1]-tt[0]))   # assume uniform spacing
    Fyy = abs(numpy.fft.fft(yy))
    guess_freq = abs(ff[numpy.argmax(Fyy[1:])+1])   # excluding the zero frequency "peak", which is related to offset
    guess_amp = numpy.std(yy) * 2.**0.5
    guess_offset = numpy.mean(yy)
    guess = numpy.array([guess_amp, 2.*numpy.pi*guess_freq, 0., guess_offset])

    def sinfunc(t, A, w, p, c):  return A * numpy.sin(w*t + p) + c
    popt, pcov = scipy.optimize.curve_fit(sinfunc, tt, yy, p0=guess)
    A, w, p, c = popt
    f = w/(2.*numpy.pi)
    fitfunc = lambda t: A * numpy.sin(w*t + p) + c
    return {"amp": A, "omega": w, "phase": p, "offset": c, "freq": f, "period": 1./f, "fitfunc": fitfunc, "maxcov": numpy.max(pcov), "rawres": (guess,popt,pcov)}

def draw_plt(tt,yy):
    N, amp, omega, phase, offset, noise = 500, 1., 2., .5, 4., 3
    #N, amp, omega, phase, offset, noise = 50, 1., .4, .5, 4., .2
    #N, amp, omega, phase, offset, noise = 200, 1., 20, .5, 4., 1
    tt = numpy.linspace(0, 10, N)
    tt2 = numpy.linspace(0, 10, 10*N)
    yy = amp*numpy.sin(omega*tt + phase) + offset
    yynoise = yy + noise*(numpy.random.random(len(tt))-0.5)

    res = fit_sin(tt, yynoise)
    print( "Amplitude=%(amp)s, Angular freq.=%(omega)s, phase=%(phase)s, offset=%(offset)s, Max. Cov.=%(maxcov)s" % res )

    plt.plot(tt, yy, "-k", label="y", linewidth=2)
    plt.plot(tt, yynoise, "ok", label="y with noise")
    plt.plot(tt2, res["fitfunc"](tt2), "r-", label="y fit curve", linewidth=2)
    plt.legend(loc="best")
    plt.show()
#---the end of calculate

def align_phase(d):
    '''
    :param d: serials or list to align to 0-2pi
    :return:
    '''
    for i,j in enumerate(d):
        while d[i] <0:
            d[i]+=2*np.pi
        while d[i]>2*np.pi:
            d[i]-=2*np.pi
    return d

def pe_cal(c1,c2):
    '''
    This is to receive the two 3*10 data series and calculate the phase of each 1 to 10GHz values and PE, SKEW
    :param c1: XI or YI(C2,C6)
    :param c2: XQ or YQ(C3,C7)
    :return: PE,SKEW,new phase array
    '''
    new=(align_phase(c2)-align_phase(c1))*360/2/np.pi
    print(new)
    for i,j in enumerate(new):
        while new[i]<0:
            new[i]+=180
        while new[i]>360:
            new[i]-=180
    for i,j in enumerate(new):
        while new[i]-min(new)>170:
            new[i]-=180
    #Old method need to align twice and two loops
    # for i,j in enumerate(new):
    #     t=0.0
    #     if i%3==0:
    #         t=min(new[i:i+3])
    #         for z in range(i,i+3):
    #             while new[z]-t>170:
    #                 new[z]-=180
    print(new)
    re=[0.0]*10
    fre=[i for i in range(1,11)]
    for i in range(10):
        re[i]=np.average(new[3*i:3*(i+1)])
    print(re)
    pe,skew,phase_new=fit_phase_diff(fre,re)
    return (pe,skew,re)

def deal_with_data(data):
    '''
    calculate the raw data
    :param data: 3*10*4 data list
    :return:[pe_x,pe_y,skew_x,skew_y]
    '''
    #deal with the data and return the PE result and plot the curve
    #transfer data to DataFrame format
    df=pd.DataFrame(data)
    phase_all=pd.DataFrame(columns=('C2','C3','C6','C7'))
    for i in range(10):
        fre=i+1
        for j in range(3):
            #C2,C3,C6,C7 equal to columns 0,1,2,3
            idx=i*3+j
            C2_t=np.array(df.iloc[idx,0][0])*1e12
            C3_t=np.array(df.iloc[idx,1][0])*1e12
            C6_t=np.array(df.iloc[idx,2][0])*1e12
            C7_t=np.array(df.iloc[idx,3][0])*1e12
            C2_v=np.array(df.iloc[idx,0][1])*1e3
            C3_v=np.array(df.iloc[idx,1][1])*1e3
            C6_v=np.array(df.iloc[idx,2][1])*1e3
            C7_v=np.array(df.iloc[idx,3][1])*1e3
            phase_all.loc[idx]=[fit_sin(C2_t,C2_v)['phase'],
                                fit_sin(C3_t,C3_v)['phase'],
                                fit_sin(C6_t,C6_v)['phase'],
                                fit_sin(C7_t,C7_v)['phase']]
    print(phase_all)
    c2=phase_all.iloc[:,0]
    c3=phase_all.iloc[:,1]
    c6=phase_all.iloc[:,2]
    c7=phase_all.iloc[:,3]
    pe_x,skew_x,new_pahseX=pe_cal(c2,c3)
    pe_y,skew_y,new_pahseY=pe_cal(c7,c6)
    return [pe_x,pe_y,skew_x,skew_y,new_pahseX,new_pahseY]


def fold_2pi(d):
    if d[0] > 180:
        d = d - 360
    if d[0] < -180:
        d = d + 360
    return d

def fit_phase_diff(freq, phase):
    """
    Fits a line to Phase_difference Vs Frequency
    to find the slope (-> skew or time delay in ps),
    and phase error (deviation from 90 degrees at 0 GHz)
    """
    mb = np.polyfit(freq, phase, 1)
    phase_new = np.polyval(mb, freq)
    #pe = mb[1]
    #skew = mb[0] / 360 * 1e12
    pe = mb[1]-90
    skew = mb[0]
    return (pe, skew, phase_new)

def fit_phase_diff_out(freq, phase, verbose=False):
    """
    Performs a linear fit while disregarding outliers until
    either a maximum number of iterations is reached or the
    slope does not change more than a threshold.
    """
    x = np.array(freq)
    y = np.array(phase)
    n_iter = 5
    change_thr = 5e-2  # %
    for i in range(n_iter):
        poly1 = np.polyfit(x, y, 1)
        fit1 = np.polyval(poly1, x)
        idx = np.argmax(np.abs(y-fit1))
        y = np.delete(y, idx)
        x = np.delete(x, idx)
        poly2 = np.polyfit(x, y, 1)
        change = np.abs(poly1[0]-poly2[0])/np.abs(poly1[0])
        if change < change_thr:
            if verbose:
                logging.info("No significant changes after {} steps".format(i))
            break
        if i >= n_iter and verbose:
            logging.info("Max number of iterations reached.")
    phase_new = np.polyval(poly2, freq)
    pe = poly2[1]
    skew = poly2[0] / 360 * 1e12
    return (pe, skew, phase_new)

# ADC_SR = 80e9  # ADC sampling rate = 80GHz need to transfer to the scope sample rate
# AOM_SHIFT = 0
#
# dt = 1.25e-11
# slen = 2 ** 19
# t = np.arange(0, slen) * dt
#
# # frequency vector used for filter construction in Frequency Domain
# f_vec = np.arange(-slen / 2, slen / 2, 1.0)
# f_vec = (f_vec / slen) / dt
# f_vec = np.fft.fftshift(f_vec)
# Gauss_BandWidth = 1e6
#---the code below is IMRA calculation codes

def measurephase_f(c1, c2, frq, dtt):
    """
    Calculate relative phase between two time-domain traces.
    """
    t = np.arange(0, len(c1)) * dtt
    s = np.exp(-1j * frq * 2 * np.pi * t)
    a1 = np.sum(c1 * s)
    a2 = np.sum(c2 * s)
    ang = np.angle(a1 / a2)
    return ang / np.pi * 180, 0


def normalize_trace(data):
    """
    Perform unit-power normalization.
    The DC component is first removed.
    """
    data = data.astype(float)
    data = data - np.mean(data)
    data = data / np.sqrt(np.mean(data * data))
    return data

def calculate_all(freq, data_list):
    '''
    calculate the PE,Skew of X and Y
    :return:
    '''
    xi=[];xq=[];yi=[];yq=[];
    phase_xiq = []
    phase_yiq = []
    phase_xiq2 = []
    phase_yiq2 = []
    fre_lo=freq#[i for i in range(1,11)] #from 1GHz to 10GHz
    #This is to check amplitude, make sure TIA is on
    for k in range(0, 4):
        d_tmp = data_list[k]
    d_pp = np.max(d_tmp) - np.min(d_tmp)
    if d_pp < 100:
        return False, 'The amplitude of at least one channel is too small, modulated at {}GHz'.format(freq*1e-9)
    for i in range(1,11):
        xi_fft = np.fft.fft(normalize_trace(data_list[i-1][0]))
        xq_fft = np.fft.fft(normalize_trace(data_list[i-1][1]))
        yi_fft = np.fft.fft(normalize_trace(data_list[i-1][2]))
        yq_fft = np.fft.fft(normalize_trace(data_list[i-1][3]))

        xi.append(xi_fft)
        xq.append(xq_fft)
        yi.append(yi_fft)
        yq.append(yq_fft)

        phase_xiq.append(measurephase_f(xi[i-1], xq[i-1], i, dt)[0] / 180 * np.pi)
        phase_yiq.append(measurephase_f(yi[i-1], yq[i-1], i, dt)[0] / 180 * np.pi)

    phase_xiq = fold_2pi(np.unwrap(np.array(phase_xiq)) / np.pi * 180)
    phase_yiq = fold_2pi(np.unwrap(np.array(phase_yiq)) / np.pi * 180)

    pe_x, skew_x, phase_newx = fit_phase_diff(fre_lo, phase_xiq)
    pe_y, skew_y, phase_newy = fit_phase_diff(fre_lo, phase_yiq)


    return (pe_x,pe_y,skew_x,skew_y)
